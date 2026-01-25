"""
Shared shape utility functions for content and icon tools.
"""


def get_shape_and_geometry(slide, shape_id=None, shape_name=None):
    """Find a shape and return its geometry.

    Args:
        slide: The slide object to search in
        shape_id: ID of the shape to find (optional)
        shape_name: Name of the shape to find (optional)

    Returns:
        Tuple of (shape, geometry_dict) or (None, error_message)
        geometry_dict contains: left, top, width, height (in inches)
    """
    shape = None
    if shape_id is not None:
        for s in slide.shapes:
            if s.shape_id == shape_id:
                shape = s
                break
        if not shape:
            return None, f"Shape with ID {shape_id} not found"
    elif shape_name is not None:
        for s in slide.shapes:
            if s.name == shape_name:
                shape = s
                break
        if not shape:
            return None, f"Shape named '{shape_name}' not found"
    else:
        return None, "Must provide either shape_id or shape_name"

    geometry = {
        'left': shape.left.inches,
        'top': shape.top.inches,
        'width': shape.width.inches,
        'height': shape.height.inches,
    }
    return shape, geometry


def delete_shape(shape):
    """Delete a shape from its slide."""
    sp = shape._element
    sp.getparent().remove(sp)
