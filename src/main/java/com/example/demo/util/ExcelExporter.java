package com.example.demo.util;

import com.example.demo.model.Employee;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

public class ExcelExporter {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<Employee> employeeList;

    public ExcelExporter(List<Employee> employees) {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Report");
        this.employeeList = employees;

    }

    private void writeHeaderRow() {
        Row row = sheet.createRow(0);
        CellStyle cellStyle = workbook.createCellStyle();
        XSSFFont fontCellHeader = workbook.createFont();
        fontCellHeader.setBold(true);
        fontCellHeader.setFontName("Times New Roman");
        fontCellHeader.setFontHeight(12);
        cellStyle.setFont(fontCellHeader);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        Cell cell = row.createCell(0);
        cell.setCellValue("STT");
//        sheet.addMergedRegion(new CellRangeAddress(0,1,0,0));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(1);
        cell.setCellValue("Đối tác thu hộ");
//        sheet.addMergedRegion(new CellRangeAddress(0,1,1,1));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(2);
        cell.setCellValue("Mã gạch nợ");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(3);
        cell.setCellValue("Kết quả trả về");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(4);
        cell.setCellValue("Số tiền");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(5);
        cell.setCellValue("Trạng thái");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(6);
        cell.setCellValue("Thời gian");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(7);
        cell.setCellValue("Nhà cung cấp");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(8);
        cell.setCellValue("Dịch vụ");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(9);
        cell.setCellValue("Mã giao dịch Merchant");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(10);
        cell.setCellValue("Mã giao dịch Partner");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(11);
        cell.setCellValue("Mã khách hàng");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(12);
        cell.setCellValue("Mã tỉnh");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(13);
        cell.setCellValue("Merchant");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(14);
        cell.setCellValue("Chi tiết giao dịch");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(15);
        cell.setCellValue("Phí");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(16);
        cell.setCellValue("Thông tin thanh toán");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(17);
        cell.setCellValue("Loại gói cước");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(18);
        cell.setCellValue("Nội dung");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(19);
        cell.setCellValue("Tính phí thu hộ");
        cell.setCellStyle(cellStyle);

    }

    private void writeDataRow() {
        CellStyle cellStyleColumn = workbook.createCellStyle();
        XSSFFont fontCellColumn = workbook.createFont();
        fontCellColumn.setFontName("Times New Roman");
        fontCellColumn.setFontHeight(12);
        cellStyleColumn.setFont(fontCellColumn);
        cellStyleColumn.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleColumn.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleColumn.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleColumn.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleColumn.setBorderLeft(HSSFCellStyle.BORDER_THIN);


        int rowCount = 1;


        for (Employee employee : employeeList) {
            Row row = sheet.createRow(rowCount++);
            Cell cell = row.createCell(0);
            cell.setCellValue(employee.getId());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(0);

            cell = row.createCell(1);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(1);

            cell = row.createCell(2);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(2);

            cell = row.createCell(3);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(3);

            cell = row.createCell(4);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(4);

            cell = row.createCell(5);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(5);

            cell = row.createCell(6);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(6);

            cell = row.createCell(7);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(7);

            cell = row.createCell(8);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(8);

            cell = row.createCell(9);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(9);

            cell = row.createCell(10);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(10);

            cell = row.createCell(11);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(11);

            cell = row.createCell(12);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(12);

            cell = row.createCell(13);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(13);

            cell = row.createCell(14);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(14);

            cell = row.createCell(15);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(15);

            cell = row.createCell(16);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(16);

            cell = row.createCell(17);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(17);

            cell = row.createCell(18);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(18);

            cell = row.createCell(19);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(19);
        }
    }

    public void export(HttpServletResponse response) throws IOException {
        writeHeaderRow();
        writeDataRow();

        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        outputStream.close();
    }


}
