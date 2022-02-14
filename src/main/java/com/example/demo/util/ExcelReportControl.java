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

public class ExcelReportControl {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<Employee> employeeList;

    public ExcelReportControl(List<Employee> employees) {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("ReportControl");
        this.employeeList = employees;

    }

    private void writeHeaderRow() {

        // Style header 1
        CellStyle cellStyle = workbook.createCellStyle();
        XSSFFont fontCellHeader = workbook.createFont();
        fontCellHeader.setBold(true);
        fontCellHeader.setFontName("Times New Roman");
        fontCellHeader.setFontHeight(11);
        cellStyle.setFont(fontCellHeader);
        cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        // Style header 2
        CellStyle cellStyle2 = workbook.createCellStyle();
        XSSFFont fontCellHeader1 = workbook.createFont();
        fontCellHeader1.setBold(true);
        fontCellHeader1.setFontName("Times New Roman");
        fontCellHeader1.setFontHeight(11);
        cellStyle2.setFont(fontCellHeader1);
        cellStyle2.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle2.setWrapText(true);

        // Size ô
        sheet.setColumnWidth(0, 10000);
        sheet.setColumnWidth(2, 4000);
        sheet.setColumnWidth(6, 4000);
        sheet.setColumnWidth(7, 4000);
        sheet.setColumnWidth(8, 4000);
        sheet.setColumnWidth(9, 4000);
        sheet.setColumnWidth(10, 4000);



        //Row 1
        final Row row = sheet.createRow(2);
        final Cell cell = row.createCell(0);
        cell.setCellValue("Báo cáo chi tiết đối soát dịch vụ bán gói");
        cell.setCellStyle(cellStyle);

        final Cell cell2 = row.createCell(1);
        cell2.setCellStyle(cellStyle);

        final Cell cell3 = row.createCell(2);
        cell3.setCellStyle(cellStyle);

        final Cell cell5 = row.createCell(4);
        cell5.setCellValue("Tên đối tác");
        cell5.setCellStyle(cellStyle);

        final Cell cell6 = row.createCell(5);
        cell6.setCellStyle(cellStyle);

        //Row 2
        final Row row2 = sheet.createRow(3);
        final Cell cell1Row2 = row2.createCell(0);
        cell1Row2.setCellValue("Loại");
        cell1Row2.setCellStyle(cellStyle);

        final Cell cell2Row2 = row2.createCell(1);
        cell2Row2.setCellValue("SLGD");
        cell2Row2.setCellStyle(cellStyle);

        final Cell cell3Row2 = row2.createCell(2);
        cell3Row2.setCellValue("Tổng tiền");
        cell3Row2.setCellStyle(cellStyle);

        //Row 3
        final Row row3 = sheet.createRow(4);
        final Cell cell1Row3 = row3.createCell(0);
        cell1Row3.setCellValue("Khớp (00)");
        cell1Row3.setCellStyle(cellStyle);

        final Cell cell2Row3 = row3.createCell(1);
        cell2Row3.setCellStyle(cellStyle);

        final Cell cell3Row3 = row3.createCell(2);
        cell3Row3.setCellStyle(cellStyle);

        //Row 4
        final Row row4 = sheet.createRow(5);
        final Cell cell1Row4 = row4.createCell(0);
        cell1Row4.setCellValue("VNPT Pay có, đối tác không có (01)");
        cell1Row4.setCellStyle(cellStyle);

        final Cell cell2Row4 = row4.createCell(1);
        cell2Row4.setCellStyle(cellStyle);

        final Cell cell3Row4 = row4.createCell(2);
        cell3Row4.setCellStyle(cellStyle);

        //Row 5
        final Row row5 = sheet.createRow(6);
        final Cell cell1Row5 = row5.createCell(0);
        cell1Row5.setCellValue("VNPT Pay không có, đối tác có (02)");
        cell1Row5.setCellStyle(cellStyle);

        final Cell cell2Row5 = row5.createCell(1);
        cell2Row5.setCellStyle(cellStyle);

        final Cell cell3Row5 = row5.createCell(2);
        cell3Row5.setCellStyle(cellStyle);

        //Row 6
        final Row row6 = sheet.createRow(7);
        final Cell cell1Row6 = row6.createCell(0);
        cell1Row6.setCellValue("Lệch (03)");
        cell1Row6.setCellStyle(cellStyle);

        final Cell cell2Row6 = row6.createCell(1);
        cell2Row6.setCellStyle(cellStyle);

        final Cell cell3Row6 = row6.createCell(2);
        cell3Row6.setCellStyle(cellStyle);

        //Row 7
        final Row row7 = sheet.createRow(8);
        final Cell cell1Row7 = row7.createCell(0);
        cell1Row7.setCellValue("Chi tiết");
        cell1Row7.setCellStyle(cellStyle);

        //Row 8

        final Row row8 = sheet.createRow(9);
        final Cell cell1Row8 = row8.createCell(0);
        cell1Row8.setCellValue("STT");
        cell1Row8.setCellStyle(cellStyle2);

        final Cell cell2Row8 = row8.createCell(1);
        cell2Row8.setCellValue("Mã đối chiếu");
        cell2Row8.setCellStyle(cellStyle2);

        final Cell cell3Row8 = row8.createCell(2);
        cell3Row8.setCellValue("Mã giao dịch VNPTPAY");
        cell3Row8.setCellStyle(cellStyle2);

        final Cell cell4Row8 = row8.createCell(3);
        cell4Row8.setCellValue("Khách hàng");
        cell4Row8.setCellStyle(cellStyle2);

        final Cell cell5Row8 = row8.createCell(4);
        cell5Row8.setCellValue("Nội dung");
        cell5Row8.setCellStyle(cellStyle2);

        final Cell cell6Row8 = row8.createCell(5);
        cell6Row8.setCellValue("Số tiền đối soát");
        cell6Row8.setCellStyle(cellStyle2);

        final Cell cell7Row8 = row8.createCell(6);
        cell7Row8.setCellValue("Thời gian giao dịch");
        cell7Row8.setCellStyle(cellStyle2);

        final Cell cell8Row8 = row8.createCell(7);
        cell8Row8.setCellValue("Trạng thái đối soát");
        cell8Row8.setCellStyle(cellStyle2);

        final Cell cell9Row8 = row8.createCell(8);
        cell9Row8.setCellValue("Mã xác nhận");
        cell9Row8.setCellStyle(cellStyle2);

        final Cell cell10Row8 = row8.createCell(9);
        cell10Row8.setCellValue("Mô tả xác nhận");
        cell10Row8.setCellStyle(cellStyle2);

        final Cell cell11Row8 = row8.createCell(10);
        cell11Row8.setCellValue("Chú thích");
        cell11Row8.setCellStyle(cellStyle2);


        //Row 9
        final Row row9 = sheet.createRow(10);
        int cellRow9 = 0;
        while(cellRow9<=10){
            final Cell cell1Row9 = row9.createCell(cellRow9);
            cell1Row9.setCellStyle(cellStyle2);
            cellRow9++;
        }


        int firstCol = 0;
        while(firstCol<= 10){
            final int cellrange1 = sheet.addMergedRegion( new CellRangeAddress(9,10,firstCol,firstCol));
            firstCol++;
        }
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
//
//
        int rowCount = 11;
//
//
        for (Employee employee : employeeList) {
            Row row = sheet.createRow(rowCount++);
            Cell cell = row.createCell(0);
            cell.setCellValue(employee.getId());
            cell.setCellStyle(cellStyleColumn);

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

        }
    }

    public void exportExcelReportControl(HttpServletResponse response) throws IOException {
        writeHeaderRow();
        writeDataRow();

        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        outputStream.close();
    }


}
