import re

def sanitize_filename(name: str) -> str:
    """
    Remove or replace characters illegal in filenames.
    """
    return re.sub(r'[\\/*?:"<>|]', "_", str(name))

def extract_short_name(full_name: str, length: int = 12) -> str:
    """
    Return the first `length` characters of a name, uppercase and stripped.
    """
    return (str(full_name)[:length]).strip().upper()

def truncate_middle(s: str, max_len: int = 30) -> str:
    """
    If `s` is longer than `max_len`, shorten it by replacing the middle with '…'.
    """
    s = str(s)
    if len(s) <= max_len:
        return s
    left_len = max_len // 2
    # Remaining for the right, after 1 char for '…'
    right_len = max_len - left_len - 1
    return s[:left_len] + '…' + s[-right_len:]
    
