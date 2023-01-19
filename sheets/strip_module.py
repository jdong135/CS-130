import decimal


def is_number(string):
    return string.isnumeric() or (string.replace('.', '', 1).isdigit() and string.count('.') < 2)


def strip_zeros(string):
    return string.rstrip('0').rstrip('.') if '.' in string else string


def strip_evaluation(eval):
    # given evaluation from lark, return value
    if type(eval) == str and is_number(eval):
        contents = strip_zeros(eval)
        return decimal.Decimal(contents)
    elif type(eval) == decimal.Decimal:
        stripped = strip_zeros(str(eval))
        return decimal.Decimal(stripped)
    else:
        return eval
