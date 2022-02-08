package com.example.demo.util;

import com.example.demo.model.Employee;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
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
        fontCellHeader.setFontName("TimeNewRoman");
        fontCellHeader.setFontHeight(12);
        cellStyle.setFont(fontCellHeader);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
//        cellStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
//        cellStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
//        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);

        Cell cell = row.createCell(0);
        cell.setCellValue("STT");
//        sheet.addMergedRegion(new CellRangeAddress(0,1,0,0));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(1);
        cell.setCellValue("Mã đối chiếu");
//        sheet.addMergedRegion(new CellRangeAddress(0,1,1,1));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(2);
        cell.setCellValue("Mã giao dịch VNPTPAY");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(3);
        cell.setCellValue("Khách hàng");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(4);
        cell.setCellValue("Nội dung");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(5);
        cell.setCellValue("Số tiền đối soát");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(6);
        cell.setCellValue("Thời gian giao dịch");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(7);
        cell.setCellValue("Trạng thái đối soát");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(8);
        cell.setCellValue("Mã xác nhận");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(9);
        cell.setCellValue("Mô tả xác nhận");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(10);
        cell.setCellValue("Chú thích");
        cell.setCellStyle(cellStyle);

    }

    private void writeDataRow() {
        CellStyle cellStyleColumn = workbook.createCellStyle();
        XSSFFont fontCellColumn = workbook.createFont();
        fontCellColumn.setFontName("TimeNewRoman");
        fontCellColumn.setFontHeight(12);
        cellStyleColumn.setFont(fontCellColumn);
        cellStyleColumn.setAlignment(CellStyle.ALIGN_CENTER);
//        cellStyleColumn.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
//        cellStyleColumn.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
//        cellStyleColumn.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
//        cellStyleColumn.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);


        int rowCount = 1;


        for(Employee employee :employeeList){
            Row row =  sheet.createRow(rowCount++);
            Cell cell  = row.createCell(0);
            cell.setCellValue(employee.getId());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(0);

            cell  = row.createCell(1);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(1);

            cell  = row.createCell(2);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(2);

            cell  = row.createCell(3);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(3);

            cell  = row.createCell(4);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(4);

            cell  = row.createCell(5);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(5);

            cell  = row.createCell(6);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(6);

            cell  = row.createCell(7);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(7);

            cell  = row.createCell(8);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(8);

            cell  = row.createCell(9);
            cell.setCellValue(employee.getEmail());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(9);

            cell  = row.createCell(10);
            cell.setCellValue(employee.getEmployeeCode());
            cell.setCellStyle(cellStyleColumn);
            sheet.autoSizeColumn(10);
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
