"""Properties mixin: read and write core document properties."""

from __future__ import annotations

from lxml import etree

from .base import CP, DC, DCTERMS

CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CUSTOM = f"{{{CUSTOM_NS}}}"
VT = f"{{{VT_NS}}}"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

# Maps property name -> XML tag
_PROP_MAP = {
    "title": f"{DC}title",
    "creator": f"{DC}creator",
    "subject": f"{DC}subject",
    "description": f"{DC}description",
    "last_modified_by": f"{CP}lastModifiedBy",
}


class PropertiesMixin:
    """Document property operations."""

    def get_properties(self) -> dict:
        """Get core document properties (title, creator, dates, etc.)."""
        tree = self._tree("docProps/core.xml")
        if tree is None:
            return {}

        def _val(tag: str) -> str:
            el = tree.find(tag)
            return el.text if el is not None and el.text else ""

        return {
            "title": _val(f"{DC}title"),
            "creator": _val(f"{DC}creator"),
            "subject": _val(f"{DC}subject"),
            "description": _val(f"{DC}description"),
            "last_modified_by": _val(f"{CP}lastModifiedBy"),
            "revision": _val(f"{CP}revision"),
            "created": _val(f"{DCTERMS}created"),
            "modified": _val(f"{DCTERMS}modified"),
        }

    def set_properties(
        self,
        *,
        title: str | None = None,
        creator: str | None = None,
        subject: str | None = None,
        description: str | None = None,
    ) -> dict:
        """Set core document properties.

        Args:
            title: Document title.
            creator: Document author/creator.
            subject: Document subject.
            description: Document description/comments.
        """
        tree = self._require("docProps/core.xml")

        updates = {
            "title": title,
            "creator": creator,
            "subject": subject,
            "description": description,
        }

        for prop_name, value in updates.items():
            if value is None:
                continue
            tag = _PROP_MAP[prop_name]
            el = tree.find(tag)
            if el is None:
                el = etree.SubElement(tree, tag)
            el.text = value

        self._mark("docProps/core.xml")
        return self.get_properties()

    def get_custom_properties(self) -> dict:
        """Get custom document properties from docProps/custom.xml."""
        root = self._tree("docProps/custom.xml")
        if root is None:
            return {}
        result = {}
        for prop in root.findall(f"{CUSTOM}property"):
            name = prop.get("name")
            child = prop[0] if len(prop) else None
            result[name] = child.text if child is not None and child.text else ""
        return result

    def set_custom_property(self, name: str, value: str, vt_type: str = "lpwstr") -> dict:
        """Upsert a custom property by name."""
        root = self._tree("docProps/custom.xml")
        if root is None:
            root = etree.Element(
                f"{CUSTOM}Properties",
                nsmap={
                    None: CUSTOM_NS,
                    "vt": VT_NS,
                },
            )
            self._trees["docProps/custom.xml"] = root
            self._link_custom_properties_part()

        existing = next(
            (p for p in root.findall(f"{CUSTOM}property") if p.get("name") == name), None
        )
        if existing is not None:
            for child in list(existing):
                existing.remove(child)
            vt_el = etree.SubElement(existing, f"{VT}{vt_type}")
            vt_el.text = value
        else:
            pids = [int(p.get("pid", 2)) for p in root.findall(f"{CUSTOM}property")]
            next_pid = max(max(pids) + 1, 2) if pids else 2
            prop = etree.SubElement(root, f"{CUSTOM}property")
            prop.set("fmtid", FMTID)
            prop.set("pid", str(next_pid))
            prop.set("name", name)
            vt_el = etree.SubElement(prop, f"{VT}{vt_type}")
            vt_el.text = value

        self._mark("docProps/custom.xml")
        return {"name": name, "value": value, "vt_type": vt_type}

    def _link_custom_properties_part(self) -> None:
        """A new docProps/custom.xml must be listed in the package relationships
        and content types, or Word treats it as an unknown stray file."""
        from .ooxml import ensure_override

        rels = self._tree("_rels/.rels")
        if rels is None and self.workdir and (self.workdir / "_rels" / ".rels").exists():
            rels = etree.parse(str(self.workdir / "_rels" / ".rels")).getroot()
            self._trees["_rels/.rels"] = rels
        rel_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        )
        if rels is not None and not any(r.get("Type") == rel_type for r in rels):
            from .ooxml import add_relationship

            add_relationship(rels, rel_type, "docProps/custom.xml")
            self._mark("_rels/.rels")
        ct = self._tree("[Content_Types].xml")
        if ct is not None and ensure_override(
            ct,
            "/docProps/custom.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        ):
            self._mark("[Content_Types].xml")

    def delete_custom_property(self, name: str) -> dict:
        """Delete a custom property by name."""
        root = self._tree("docProps/custom.xml")
        if root is not None:
            prop = root.find(f"{CUSTOM}property[@name='{name}']")
            if prop is not None:
                root.remove(prop)
                self._mark("docProps/custom.xml")
                return {"deleted": name}
        raise ValueError(f"Custom property not found: {name}")
