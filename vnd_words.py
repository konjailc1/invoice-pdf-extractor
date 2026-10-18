"""Read a whole-đồng amount the way a Vietnamese invoice does.

    15465000 -> "Mười lăm triệu bốn trăm sáu mươi lăm nghìn đồng chẵn"

The same function is used to print the invoice and to check that the
extracted sentence matches the total.
"""

import re

_DIGITS = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
_SCALES = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"]


def amount_in_words(amount):
    """Return the Vietnamese reading of ``amount``, with a capital first letter."""
    amount = int(amount)
    if amount < 0:
        return "Âm " + amount_in_words(-amount)
    if amount == 0:
        return "Không đồng chẵn"

    groups = []
    remaining = amount
    while remaining:
        groups.append(remaining % 1000)
        remaining //= 1000

    parts = []
    for index in range(len(groups) - 1, -1, -1):
        group = groups[index]
        if group == 0:
            continue
        higher_exists = any(groups[higher] for higher in range(index + 1, len(groups)))
        parts.append(_read_group(group, show_zero_hundred=higher_exists and group < 100))
        if _SCALES[index]:
            parts.append(_SCALES[index])

    sentence = " ".join(parts)
    sentence = sentence[0].upper() + sentence[1:]
    if amount % 1000 == 0:
        return sentence + " đồng chẵn"
    return sentence + " đồng"


def normalize_amount_words(text):
    """Lowercase and collapse spacing so a wrapped PDF line still compares equal."""
    if not text:
        return ""
    cleaned = str(text).lower().replace("đ", "đ")
    cleaned = re.sub(r"[.,;:]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _read_group(number, show_zero_hundred):
    """Read one group of up to three digits (0..999)."""
    hundreds = number // 100
    tens = (number // 10) % 10
    ones = number % 10
    words = []

    if hundreds:
        words.append(f"{_DIGITS[hundreds]} trăm")
    elif show_zero_hundred and (tens or ones):
        words.append("không trăm")

    if tens > 1:
        words.append(f"{_DIGITS[tens]} mươi")
        if ones == 1:
            words.append("mốt")
        elif ones == 4:
            words.append("tư")
        elif ones == 5:
            words.append("lăm")
        elif ones:
            words.append(_DIGITS[ones])
    elif tens == 1:
        words.append("mười")
        if ones == 5:
            words.append("lăm")
        elif ones:
            words.append(_DIGITS[ones])
    elif ones:
        if words:
            words.append("lẻ")
        words.append(_DIGITS[ones])

    return " ".join(words)


def _self_check():
    """Fail fast if the reading of the example amount drifts."""
    expected = "Mười lăm triệu bốn trăm sáu mươi lăm nghìn đồng chẵn"
    got = amount_in_words(15_465_000)
    if got != expected:
        raise SystemExit(f"amount_in_words(15465000) = {got!r}, expected {expected!r}")
    samples = {
        0: "Không đồng chẵn",
        15: "Mười lăm đồng",
        21: "Hai mươi mốt đồng",
        1_000: "Một nghìn đồng chẵn",
        1_000_000: "Một triệu đồng chẵn",
        1_105: "Một nghìn một trăm lẻ năm đồng",
    }
    for amount, reading in samples.items():
        got = amount_in_words(amount)
        if got != reading:
            raise SystemExit(f"amount_in_words({amount}) = {got!r}, expected {reading!r}")


if __name__ == "__main__":
    _self_check()
    print("vnd_words ok")
