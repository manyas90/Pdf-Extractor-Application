def format_size(size):

    kb = 1024
    mb = kb * 1024
    gb = mb * 1024

    if size >= gb:
        return f"{size / gb:.2f} GB"

    elif size >= mb:
        return f"{size / mb:.2f} MB"

    elif size >= kb:
        return f"{size / kb:.2f} KB"

    else:
        return f"{size} Bytes"