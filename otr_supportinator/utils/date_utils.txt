from datetime import datetime, timedelta, date

def get_amazon_week(input_date):
    if isinstance(input_date, date):
        input_date = datetime.combine(input_date, datetime.min.time())
    elif not isinstance(input_date, datetime):
        raise ValueError("Input must be a date or datetime object")

    # Find the last Sunday of the previous year (start of week 1)
    year_start = datetime(input_date.year, 1, 1)
    week_1_start = year_start - timedelta(days=(year_start.weekday() + 1) % 7)
    
    if input_date < week_1_start:
        # Date is in the last week of the previous year
        week_1_start = week_1_start - timedelta(days=7)
    
    # Calculate the number of weeks since the start of week 1
    weeks = (input_date - week_1_start).days // 7 + 1
    
    # If week number is 53, change it to 1
    return 1 if weeks == 53 else weeks

def get_current_amazon_week():
    return get_amazon_week(datetime.now())


def get_amazon_year(input_date):
    if isinstance(input_date, date):
        input_date = datetime.combine(input_date, datetime.min.time())
    elif not isinstance(input_date, datetime):
        raise ValueError("Input must be a date or datetime object")

    year = input_date.year
    week = get_amazon_week(input_date)
    
    # If it's week 1 and in December, it's actually part of next year
    if week == 1 and input_date.month == 12:
        year += 1
    
    return year

def get_amazon_week_start(year, week):
    # Find the last Sunday of the previous year
    year_start = datetime(year - 1, 12, 31)
    while year_start.weekday() != 6:  # 6 is Sunday
        year_start -= timedelta(days=1)
    
    # Add the number of weeks
    return year_start + timedelta(weeks=week-1)

def get_amazon_week_end(year, week):
    week_start = get_amazon_week_start(year, week)
    return week_start + timedelta(days=6)
