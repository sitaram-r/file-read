def format_response(data):
    """Formats API response"""
    return {
        "status": "success" if "Error" not in data else "failed",
        "data": data
    }
