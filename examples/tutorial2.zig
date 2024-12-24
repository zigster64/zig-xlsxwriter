const xlsxwriter = @import("xlsxwriter");

const Expense = struct {
    item: [*:0]const u8,
    cost: f64,
    datetime: xlsxwriter.lxw_datetime,
};

var expenses = [_]Expense{
    .{ .item = "Rent", .cost = 1000.0, .datetime = .{ .year = 2013, .month = 1, .day = 13, .hour = 8, .min = 34, .sec = 65.45 } },
    .{ .item = "Gas", .cost = 100.0, .datetime = .{ .year = 2013, .month = 1, .day = 14, .hour = 12, .min = 17, .sec = 23.34 } },
    .{ .item = "Food", .cost = 300.0, .datetime = .{ .year = 2013, .month = 1, .day = 16, .hour = 4, .min = 29, .sec = 54.05 } },
    .{ .item = "Gym", .cost = 50.0, .datetime = .{ .year = 2013, .month = 1, .day = 20, .hour = 6, .min = 55, .sec = 32.16 } },
};

pub fn main() void {

    // Create a workbook and add a worksheet.
    const workbook: ?*xlsxwriter.lxw_workbook = xlsxwriter.workbook_new("out/tutorial2.xlsx");
    const worksheet: ?*xlsxwriter.lxw_worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Start from the first cell. Rows and columns are zero indexed.
    var row: u32 = 0;
    const col: u16 = 0;

    // Add a bold format to use to highlight cells.
    const bold: ?*xlsxwriter.lxw_format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    const bold_rightjustified: ?*xlsxwriter.lxw_format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold_rightjustified);
    _ = xlsxwriter.format_set_align(bold_rightjustified, xlsxwriter.LXW_ALIGN_RIGHT);

    // Add a number format for cells with money.
    const money: ?*xlsxwriter.lxw_format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_num_format(money, "$#,##0");

    // Add an Excel date format.
    const date_format: ?*xlsxwriter.lxw_format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_num_format(date_format, "mmmm d yyyy");

    // Adjust the column width.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 15, null);

    // Write some data header.
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col, "Item", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col + 1, "Date", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col + 2, "Cost", bold_rightjustified);
    row += 1;

    // Iterate over the data and write it out element by element.
    var index: usize = 0;
    for (expenses) |_| {
        _ = xlsxwriter.worksheet_write_string(worksheet, row, col, expenses[index].item, null);
        _ = xlsxwriter.worksheet_write_datetime(worksheet, row, col + 1, &expenses[index].datetime, date_format);
        _ = xlsxwriter.worksheet_write_number(worksheet, row, col + 2, expenses[index].cost, money);
        index += 1;
        row += 1;
    }

    // Write a total using a formula.
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col, "Total", bold);
    _ = xlsxwriter.worksheet_write_formula(worksheet, row, col + 2, "=SUM(C2:C5)", bold_rightjustified);
    row += 1;

    // Save the workbook and free any allocated memory.
    _ = xlsxwriter.workbook_close(workbook);
}
