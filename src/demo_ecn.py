"""Small revision-label helper used by the portfolio demo."""

REVISION_LETTERS = 'ABCDEFGHJKLMNPQRSTUVWXYZ'


def next_revision(revision):
    text = str(revision or '').strip().upper()
    if not text or len(text) > 2 or any(not 'A' <= char <= 'Z' for char in text):
        return ''
    positions = [max(i for i, letter in enumerate(REVISION_LETTERS)
                     if letter <= char) for char in text]
    position = len(positions) - 1
    while position >= 0:
        positions[position] += 1
        if positions[position] < len(REVISION_LETTERS):
            break
        positions[position] = 0
        position -= 1
    if position < 0:
        positions.insert(0, 0)
    return ''.join(REVISION_LETTERS[i] for i in positions)
