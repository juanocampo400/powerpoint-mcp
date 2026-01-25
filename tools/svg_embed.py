"""
SVG embedding utilities for recolorable icons in PowerPoint.

This module handles dual-format SVG+PNG embedding via OOXML manipulation,
allowing icons to be recolored in PowerPoint via Graphics Fill.
"""

import logging
import re
import tempfile
import os
import zipfile
from pathlib import Path
from lxml import etree

logger = logging.getLogger(__name__)

# OOXML namespaces
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'asvg': 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
}

# SVG Extension GUID for PowerPoint
SVG_EXTENSION_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"


def make_svg_recolorable(svg_content: str, fill_color: str = None) -> str:
    """Strip fill/stroke color attributes from SVG for PowerPoint recolorability.

    Removes fill and stroke attributes (except fill="none") so PowerPoint
    can apply its own fill color via Graphics Fill. Optionally applies an
    initial fill color to the SVG root element.

    Args:
        svg_content: Raw SVG content string
        fill_color: Optional hex color to apply as initial fill (e.g., "#333333")

    Returns:
        Modified SVG content with colors stripped and optional fill applied
    """
    # Remove fill="currentColor" and any explicit fill colors (but keep fill="none")
    svg_content = re.sub(r'\s*fill=["\'](?!none)[^"\']*["\']', '', svg_content)

    # Remove stroke="currentColor"
    svg_content = re.sub(r'\s*stroke=["\']currentColor["\']', '', svg_content)

    # Apply fill color to SVG root if specified (for initial display color)
    # PowerPoint's Graphics Fill can still override this for recoloring
    if fill_color:
        color_val = fill_color.lstrip('#')
        svg_content = svg_content.replace('<svg ', f'<svg fill="#{color_val}" ', 1)

    return svg_content


def generate_png_fallback(svg_content: str, color: str, size_px: int) -> bytes:
    """Generate colored PNG from SVG for backwards compatibility.

    Args:
        svg_content: Raw SVG content (with or without colors)
        color: Hex color code (e.g., "#333333")
        size_px: Output size in pixels

    Returns:
        PNG image as bytes
    """
    try:
        import cairosvg
    except ImportError:
        raise ImportError("cairosvg is required for PNG generation. Run: pip install cairosvg")

    # Apply color to SVG
    colored_svg = re.sub(
        r'(stroke|fill)=["\']currentColor["\']',
        rf'\1="{color}"',
        svg_content,
        flags=re.IGNORECASE
    )

    # Also set default fill for shapes without explicit fill
    # Add fill attribute to the root SVG element if not present
    if 'fill="' not in colored_svg.split('>')[0]:
        colored_svg = colored_svg.replace('<svg ', f'<svg fill="{color}" ', 1)

    return cairosvg.svg2png(
        bytestring=colored_svg.encode('utf-8'),
        output_width=size_px,
        output_height=size_px
    )


class SVGEmbedder:
    """Handles dual-format SVG+PNG embedding via OOXML.

    PowerPoint supports dual-format images where an SVG is embedded alongside
    a PNG fallback. The SVG is used for display and can be recolored via
    Graphics Fill, while the PNG provides compatibility.
    """

    def embed_recolorable_icon(
        self,
        slide,
        svg_content: str,
        png_bytes: bytes,
        left_inches: float,
        top_inches: float,
        size_inches: float,
        icon_name: str
    ) -> int:
        """Insert icon with SVG+PNG dual format for recolorability.

        Args:
            slide: python-pptx slide object
            svg_content: Recolorable SVG content (with optional fill color applied)
            png_bytes: PNG fallback image bytes
            left_inches: Left position in inches
            top_inches: Top position in inches
            size_inches: Size in inches (icons are square)
            icon_name: Name of the icon for logging

        Returns:
            Shape ID of the inserted icon
        """
        from pptx.util import Inches
        from pptx.opc.package import Part
        from pptx.opc.packuri import PackURI

        # Create temp file for PNG
        fd, png_path = tempfile.mkstemp(suffix='.png')
        try:
            os.write(fd, png_bytes)
            os.close(fd)

            # Add PNG as base image
            shape = slide.shapes.add_picture(
                png_path,
                Inches(left_inches),
                Inches(top_inches),
                Inches(size_inches),
                Inches(size_inches)
            )
        finally:
            try:
                os.unlink(png_path)
            except OSError:
                pass

        # Get the slide part and package for relationship management
        slide_part = slide.part
        package = slide_part.package

        # Find the next available image number for SVG
        max_num = 0
        for part in package.iter_parts():
            pn = str(part.partname)
            if '/ppt/media/image' in pn:
                match = re.search(r'image(\d+)', pn)
                if match:
                    max_num = max(max_num, int(match.group(1)))

        next_num = max_num + 1
        svg_partname = PackURI(f'/ppt/media/image{next_num}.svg')

        # Create SVG part
        svg_bytes = svg_content.encode('utf-8')
        svg_part = Part(svg_partname, 'image/svg+xml', package, svg_bytes)

        # Create relationship from slide to SVG
        svg_rid = slide_part.relate_to(
            svg_part,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
        )

        # Modify the blip XML to add SVG extension
        self._add_svg_extension_to_shape(shape, svg_rid)

        logger.debug(f"Embedded recolorable icon '{icon_name}' with SVG rId={svg_rid}")
        return shape.shape_id

    def _add_svg_extension_to_shape(self, shape, svg_rid: str):
        """Add SVG extension to the shape's blip element.

        Modifies the shape's XML to include the SVG reference:

        <a:blip r:embed="rId_png">
          <a:extLst>
            <a:ext uri="{96DAC541-...}">
              <asvg:svgBlip r:embed="rId_svg"/>
            </a:ext>
          </a:extLst>
        </a:blip>
        """
        # Get the picture element
        pic = shape._element

        # Find the blip element
        blip = pic.find('.//a:blip', namespaces=NAMESPACES)
        if blip is None:
            logger.warning("Could not find blip element in shape")
            return

        # Create or get extLst
        extLst = blip.find('a:extLst', namespaces=NAMESPACES)
        if extLst is None:
            extLst = etree.SubElement(blip, '{%s}extLst' % NAMESPACES['a'])

        # Create the SVG extension
        ext = etree.SubElement(extLst, '{%s}ext' % NAMESPACES['a'])
        ext.set('uri', SVG_EXTENSION_URI)

        # Create the svgBlip element
        svgBlip = etree.SubElement(ext, '{%s}svgBlip' % NAMESPACES['asvg'])
        svgBlip.set('{%s}embed' % NAMESPACES['r'], svg_rid)


def ensure_svg_content_type(pptx_path: str):
    """Add SVG MIME type to [Content_Types].xml if missing.

    PowerPoint requires the SVG content type to be registered for
    proper handling of embedded SVG files.

    Args:
        pptx_path: Path to the saved .pptx file
    """
    content_types_entry = '<Default Extension="svg" ContentType="image/svg+xml"/>'

    # Open the pptx as a zip file
    with zipfile.ZipFile(pptx_path, 'r') as zf:
        content_types = zf.read('[Content_Types].xml').decode('utf-8')

    # Check if SVG content type already exists
    if 'Extension="svg"' in content_types or "Extension='svg'" in content_types:
        logger.debug("SVG content type already present")
        return

    # Add SVG content type before the closing Types tag
    if '</Types>' in content_types:
        content_types = content_types.replace(
            '</Types>',
            f'  {content_types_entry}\n</Types>'
        )
    else:
        logger.warning("Could not find </Types> in [Content_Types].xml")
        return

    # Write back to the zip file
    # We need to recreate the zip with the modified content
    temp_path = pptx_path + '.tmp'

    with zipfile.ZipFile(pptx_path, 'r') as zf_read:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_write:
            for item in zf_read.namelist():
                if item == '[Content_Types].xml':
                    zf_write.writestr(item, content_types.encode('utf-8'))
                else:
                    zf_write.writestr(item, zf_read.read(item))

    # Replace original with modified
    os.replace(temp_path, pptx_path)
    logger.debug("Added SVG content type to [Content_Types].xml")
