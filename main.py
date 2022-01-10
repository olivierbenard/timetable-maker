import xlsxwriter
import datetime
from functions import utils


colors = {"title": "03045e", "weekend": "caf0f8", "feiertag": "#ECEEF0"}


def main():

    current_date = datetime.datetime.now()
    current_year = current_date.year
    current_month = current_date.month

    days = utils.list_days(current_year)

    workbook = xlsxwriter.Workbook(f"timetable_{current_year}.xlsx")
    worksheet = workbook.add_worksheet()

    format_title = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": colors["title"],
            "align": "center",
            "valign": "vcenter",
        }
    )
    format_hsep = workbook.add_format(
        {"bg_color": "white", "align": "center", "valign": "vcenter"}
    )
    format_vsep = workbook.add_format(
        {"bg_color": "white", "align": "center", "valign": "vcenter"}
    )
    format_day = workbook.add_format(
        {
            "bg_color": "white",
            "align": "center",
            "valign": "right",
            "bottom": 1,
            "top": 1,
        }
    )
    format_number = workbook.add_format(
        {
            "bg_color": "white",
            "align": "center",
            "valign": "vcenter",
            "bottom": 1,
            "top": 1,
        }
    )
    format_empty_cell = workbook.add_format(
        {
            "bg_color": "white",
            "align": "center",
            "valign": "vcenter",
            "bottom": 1,
            "top": 1,
        }
    )
    format_day_we = workbook.add_format(
        {
            "bg_color": colors["weekend"],
            "align": "center",
            "valign": "right",
            "bottom": 1,
            "top": 1,
        }
    )
    format_number_we = workbook.add_format(
        {
            "bg_color": colors["weekend"],
            "align": "center",
            "valign": "vcenter",
            "bottom": 1,
            "top": 1,
        }
    )
    format_empty_cell_we = workbook.add_format(
        {
            "bg_color": colors["weekend"],
            "align": "center",
            "valign": "vcenter",
            "bottom": 1,
            "top": 1,
        }
    )
    format_day_feiertag = workbook.add_format(
        {
            "bg_color": colors["feiertag"],
            "align": "center",
            "valign": "right",
            "bottom": 1,
            "top": 1,
        }
    )
    format_number_feiertag = workbook.add_format(
        {
            "bg_color": colors["feiertag"],
            "align": "center",
            "valign": "vcenter",
            "bottom": 1,
            "top": 1,
        }
    )
    format_empty_cell_feiertag = workbook.add_format(
        {
            "bg_color": colors["feiertag"],
            "align": "center",
            "valign": "vcenter",
            "bottom": 1,
            "top": 1,
        }
    )

    row, col = 1, 0

    # horizontal separator
    worksheet.set_row(1, 5)
    for j in range(0, 70):
        worksheet.write(1, 1 + j, "", format_hsep)

    for i in range(12):
        # vertical separator
        worksheet.set_column(5 + 6 * i, 5 + 6 * i, width=1)
        # weekday separator
        worksheet.set_column(6 * i, 6 * i, width=5)
        # number separator
        worksheet.set_column(1 + 6 * i, 1 + 6 * i, width=3)

    worksheet.write(1, 0, "", format_vsep)

    for day in days:

        # if first month to display
        if (col == 0) and (row == 1):
            worksheet.merge_range(
                0, col, 0, col + 4, utils.datetime_to_month_str(day), format_title
            )
            row += 1

        # if enterring the next month
        if day.month != current_month:

            current_month = day.month
            col += 6
            worksheet.merge_range(
                0, col, 0, col + 4, utils.datetime_to_month_str(day), format_title
            )
            row = 2

        if utils.is_bank_holiday(day) and not (
            day.weekday() == 5 or day.weekday() == 6
        ):
            format_day_temp = format_day_feiertag
            format_number_temp = format_number_feiertag
            format_empty_cell_temp = format_empty_cell_feiertag
        elif day.weekday() == 5 or day.weekday() == 6:
            format_day_temp = format_day_we
            format_number_temp = format_number_we
            format_empty_cell_temp = format_empty_cell_we
        else:
            format_day_temp = format_day
            format_number_temp = format_number
            format_empty_cell_temp = format_empty_cell

        worksheet.write(row, col, utils.datetime_to_weekday_str(day), format_day_temp)
        worksheet.write(row, col + 1, str(day.day).zfill(2), format_number_temp)
        worksheet.write(row, col + 2, "", format_empty_cell_temp)
        worksheet.write(row, col + 3, "", format_empty_cell_temp)
        worksheet.write(row, col + 4, "", format_empty_cell_temp)
        worksheet.write(row, col + 5, "", format_vsep)

        row += 1

    workbook.close()


if __name__ == "__main__":
    main()
