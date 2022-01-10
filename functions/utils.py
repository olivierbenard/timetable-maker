import datetime


def list_days(year):

    """
    Returns the list of dates from January, 1 to December, 31 of the specified year.
    """

    start_date = datetime.datetime(year, 1, 1)
    end_date = datetime.datetime(year, 12, 31)

    nb_days = (end_date - start_date).days + 1

    return (start_date + datetime.timedelta(days=i) for i in range(nb_days))


def datetime_to_month_str(dt):

    """
    Returns a %B-%Y formated string from datetime object.
    """

    return dt.strftime("%B-%Y")


def datetime_to_weekday_str(dt):

    """
    Returns a %a formated string from datetime object.
    """

    return dt.strftime("%a")


def is_bank_holiday(dt):

    """
    Returns True if datetime corresponds to a german feiertag
    """

    return dt.strftime("%d/%m") in (
        "01/01",
        "06/01",
        "15/04",
        "17/04",
        "18/04",
        "01/05",
        "26/05",
        "06/06",
        "16/06",
        "06/09",
        "03/10",
        "01/11",
        "24/12",
        "25/12",
        "26/12",
        "31/12",
    )


def int_to_month(n):
    return {
        1: "janvier",
        2: "février",
        3: "mars",
        4: "avril",
        5: "mai",
        6: "juin",
        7: "juillet",
        8: "août",
        9: "septembre",
        10: "octobre",
        11: "novembre",
        12: "décember",
    }[n]


if __name__ == "__main__":
    pass
