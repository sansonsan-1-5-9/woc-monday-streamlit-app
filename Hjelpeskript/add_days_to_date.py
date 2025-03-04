from datetime import datetime, timedelta


def calculate_easter(year):
    """Beregn datoen for første påskedag basert på algoritmen av Meeus/Jones/Butcher."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return datetime(year, month, day)


def norwegian_holidays(year):
    """Returnerer en liste over røde dager i Norge for et gitt år."""
    # Faste røde dager
    holidays = [
        datetime(year, 1, 1),  # Nyttårsdag
        datetime(year, 5, 1),  # Arbeidernes dag
        datetime(year, 5, 17),  # Grunnlovsdagen
        datetime(year, 12, 25),  # Første juledag
        datetime(year, 12, 26)  # Andre juledag
    ]

    # Bevegelige røde dager
    easter_sunday = calculate_easter(year)
    holidays += [
        easter_sunday - timedelta(days=3),  # Skjærtorsdag
        easter_sunday - timedelta(days=2),  # Langfredag
        easter_sunday + timedelta(days=1),  # Andre påskedag
        easter_sunday + timedelta(days=39),  # Kristi himmelfartsdag
        easter_sunday + timedelta(days=49),  # Første pinsedag
        easter_sunday + timedelta(days=50)  # Andre pinsedag
    ]
    return set(holidays)  # Returner som et sett for rask oppslag


def add_working_days_with_holidays(start_date: str, days_to_add: int) -> str:
    # Convert ISO date format if necessary
    start_date = start_date.split("T")[0]  # Extract only "YYYY-MM-DD"

    # Konverterer startdato til datetime-objekt
    start = datetime.strptime(start_date, "%Y-%m-%d")
    current_date = start
    days_added = 0

    # Hent helligdager for året det gjelder
    holidays = norwegian_holidays(start.year)

    # Iterer til vi har lagt til det ønskede antallet arbeidsdager
    while days_added < days_to_add:
        current_date += timedelta(days=1)
        # Sjekker om dagen er en arbeidsdag og ikke en rød dag
        if current_date.weekday() < 5 and current_date not in holidays:
            days_added += 1

        # Hvis vi krysser til et nytt år, legg til røde dager for det nye året
        if current_date.year != start.year:
            holidays.update(norwegian_holidays(current_date.year))

    # Returnerer datoen som streng i samme format
    return current_date.strftime("%Y-%m-%d")


# date = "2025-01-29T07:42:32.347Z"
#
# newdate = add_working_days_with_holidays(date,5)
#
# print(newdate)