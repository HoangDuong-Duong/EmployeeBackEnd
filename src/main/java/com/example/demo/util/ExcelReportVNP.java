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

public class ExcelReportVNP {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<Employee> employeeList;

    public ExcelReportVNP(List<Employee> employees) {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("ReportVNP");
        this.employeeList = employees;

    }

    private void writeHeaderRow() {

        CellStyle cellStyle = workbook.createCellStyle();
        XSSFFont fontCellHeader = workbook.createFont();
        fontCellHeader.setBold(true);
        fontCellHeader.setFontName("Times New Roman");
        fontCellHeader.setFontHeight(12);
        cellStyle.setFont(fontCellHeader);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle.setWrapText(true);

        CellStyle cellStyle2 = workbook.createCellStyle();
        XSSFFont fontCellHeader2 = workbook.createFont();
        fontCellHeader2.setBold(true);
        fontCellHeader2.setFontName("Times New Roman");
        fontCellHeader2.setFontHeight(14);
        cellStyle2.setFont(fontCellHeader2);
        cellStyle2.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle2.setWrapText(true);


        CellStyle cellStyle3 = workbook.createCellStyle();
        XSSFFont fontCellHeader3 = workbook.createFont();
        fontCellHeader3.setFontName("Times New Roman");
        fontCellHeader3.setFontHeight(12);
        cellStyle3.setFont(fontCellHeader3);
        cellStyle3.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle3.setWrapText(true);

        CellStyle cellStyleItalic = workbook.createCellStyle();
        XSSFFont fontCellHeader4 = workbook.createFont();
        fontCellHeader4.setFontName("Times New Roman");
        fontCellHeader4.setFontHeight(12);
        fontCellHeader4.setItalic(true);
        cellStyleItalic.setFont(fontCellHeader4);
        cellStyleItalic.setAlignment(CellStyle.ALIGN_LEFT);
        cellStyleItalic.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleItalic.setWrapText(true);


        CellStyle cellStyleValue = workbook.createCellStyle();
        XSSFFont fontCellHeader5 = workbook.createFont();
        fontCellHeader5.setBold(true);
        fontCellHeader5.setFontName("Times New Roman");
        fontCellHeader5.setFontHeight(10.5);
        cellStyleValue.setFont(fontCellHeader5);
        cellStyleValue.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleValue.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleValue.setWrapText(true);


        CellStyle cellStyleTable = workbook.createCellStyle();
        XSSFFont fontCellHeader7 = workbook.createFont();
        fontCellHeader7.setBold(true);
        fontCellHeader7.setFontName("Times New Roman");
        fontCellHeader7.setFontHeight(10.5);
        cellStyleTable.setFont(fontCellHeader7);
        cellStyleTable.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleTable.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleTable.setWrapText(true);
        cellStyleTable.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleTable.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleTable.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleTable.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        CellStyle cellStyle7 = workbook.createCellStyle();
        XSSFFont fontCellHeader8 = workbook.createFont();
        fontCellHeader8.setBold(true);
        fontCellHeader8.setItalic(true);
        fontCellHeader8.setFontName("Times New Roman");
        fontCellHeader8.setFontHeight(12);
        cellStyle7.setFont(fontCellHeader8);
        cellStyle7.setAlignment(CellStyle.ALIGN_LEFT);
        cellStyle7.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle7.setWrapText(true);

        CellStyle cellStyle8 = workbook.createCellStyle();
        XSSFFont fontCellHeader9 = workbook.createFont();
        fontCellHeader9.setItalic(true);
        fontCellHeader9.setFontName("Times New Roman");
        fontCellHeader9.setFontHeight(12);
        cellStyle8.setFont(fontCellHeader9);
        cellStyle8.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle8.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle8.setWrapText(true);




        //Size ô
        sheet.setColumnWidth(0, 2719);
        sheet.setColumnWidth(1, 8650);
        sheet.setColumnWidth(2, 3482);
        sheet.setColumnWidth(3, 3883);
        sheet.setColumnWidth(4, 4042);
        sheet.setColumnWidth(5, 4081);
        sheet.setColumnWidth(6, 4563);
        sheet.setColumnWidth(7, 4482);
        sheet.setColumnWidth(8, 5241);
        sheet.setColumnWidth(9, 4843);






        //Row1
        Row row1 = sheet.createRow(0);
        row1.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell = row1.createCell(0);
        cell.setCellValue("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM");
        final int cellrange1 = sheet.addMergedRegion( new CellRangeAddress(0,0,0,8));
        cell.setCellStyle(cellStyle);

        //Row2
        Row row2 = sheet.createRow(1);
        row2.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell1 = row2.createCell(0);
        cell1.setCellValue("Độc lập- Tự do- Hạnh phúc");
        final int cellrange2 = sheet.addMergedRegion( new CellRangeAddress(1,1,0,8));
        cell1.setCellStyle(cellStyle);

        //Row3
        Row row3 = sheet.createRow(2);
        row3.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell2 = row3.createCell(0);
        cell2.setCellValue("BIÊN BẢN ĐỐI SOÁT DÒNG TIỀN THANH TOÁN DỊCH VỤ BÁN GÓI");
        final int cellrange3 = sheet.addMergedRegion( new CellRangeAddress(2,2,0,8));
        cell2.setCellStyle(cellStyle2);

        //Row 4
        Row row4 = sheet.createRow(3);
        row4.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell3 = row4.createCell(0);
        cell3.setCellValue("Giữa …. và Tổng Công ty Truyền thông (VNPT - MEDIA)");
        final int cellrange4 = sheet.addMergedRegion( new CellRangeAddress(3,3,0,8));
        cell3.setCellStyle(cellStyle);

        //Row 5
        Row row5 = sheet.createRow(4);
        Cell cell4 = row5.createCell(0);
        cell4.setCellValue("Từ ngày 01/11/2021 đến hết ngày 30/11/2021");
        final int cellrange5 = sheet.addMergedRegion( new CellRangeAddress(4,4,0,8));
        cell4.setCellStyle(cellStyle3);

        //Row 6
        Row row6 = sheet.createRow(5);
        Cell cell5 = row6.createCell(0);
        cell5.setCellValue("- Căn cứ Hợp đồng ... giữa … (Bên A) và VNPT-Media về việc cung cấp ….. ký ngày ....");
        final int cellrange6 = sheet.addMergedRegion( new CellRangeAddress(5,5,0,8));
        cell5.setCellStyle(cellStyleItalic);

        //Row 7
        Row row7 = sheet.createRow(6);
        row7.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell6 = row7.createCell(0);
        cell6.setCellValue("- Hôm nay, ngày  tháng  năm 2021, …......... và Tổng công ty Truyền thông cùng ký xác nhận Biên bản xác nhận giao dịch dịch vụ bán gói từ ngày …/…/20...  đến hết ngày …/…/20… cụ thể như sau:");
        final int cellrange7 = sheet.addMergedRegion( new CellRangeAddress(6,6,0,9));
        cell6.setCellStyle(cellStyleItalic);

        //Row 8
        Row row8 = sheet.createRow(7);
        Cell cell7 = row8.createCell(0);
        cell7.setCellValue("Loại gói cước");
        final int cellrange8 = sheet.addMergedRegion( new CellRangeAddress(7,10,0,0));
        cell7.setCellStyle(cellStyleTable);

        Cell cell8 = row8.createCell(1);
        cell8.setCellValue("Đối tượng");
        final int cellrange9 = sheet.addMergedRegion( new CellRangeAddress(7,10,1,1));
        cell8.setCellStyle(cellStyleTable);

        Cell cell9 = row8.createCell(2);
        cell9.setCellValue("Số tiền chưa chuyển về Tài khoản ..... cuối tháng T-1");
        final int cellrange10 = sheet.addMergedRegion( new CellRangeAddress(7,10,2,2));
        cell9.setCellStyle(cellStyleTable);

        Cell cell10 = row8.createCell(3);
        cell10.setCellValue("Số phát sinh trong kỳ");
        final int cellrange11 = sheet.addMergedRegion( new CellRangeAddress(7,7,3,8));
        cell10.setCellStyle(cellStyleTable);

        Cell cell80 = row8.createCell(4);
        cell80.setCellStyle(cellStyleTable);

        Cell cell81 = row8.createCell(5);
        cell81.setCellStyle(cellStyleTable);

        Cell cell82 = row8.createCell(6);
        cell82.setCellStyle(cellStyleTable);

        Cell cell83 = row8.createCell(7);
        cell83.setCellStyle(cellStyleTable);

        Cell cell84 = row8.createCell(8);
        cell84.setCellStyle(cellStyleTable);

        Cell cell11 = row8.createCell(9);
        cell11.setCellValue("Số tiền còn chưa chuyển về Tài khoản …. ngày cuối tháng T");
        final int cellrange20 = sheet.addMergedRegion( new CellRangeAddress(7,10,9,9));
        cell11.setCellStyle(cellStyleTable);

        //Row 9
        Row row9 = sheet.createRow(8);
        Cell cell85 = row9.createCell(0);
        cell85.setCellStyle(cellStyleTable);

        Cell cell86 = row9.createCell(1);
        cell86.setCellStyle(cellStyleTable);

        Cell cell87 = row9.createCell(2);
        cell87.setCellStyle(cellStyleTable);

        Cell cell12 = row9.createCell(3);
        cell12.setCellValue("Số phát sinh tăng trong kỳ");
        final int cellrange12 = sheet.addMergedRegion( new CellRangeAddress(8,8,3,7));
        cell12.setCellStyle(cellStyleTable);

        Cell cell88 = row9.createCell(4);
        cell88.setCellStyle(cellStyleTable);

        Cell cell89 = row9.createCell(6);
        cell89.setCellStyle(cellStyleTable);


        Cell cell13 = row9.createCell(8);
        cell13.setCellValue("Số phát sinh giảm trong kỳ (Số tiền đã chuyển về Tài khoản …");
        final int cellrange13 = sheet.addMergedRegion( new CellRangeAddress(8,10,8,8));
        cell13.setCellStyle(cellStyleTable);

        Cell cell90 = row9.createCell(9);
        cell90.setCellStyle(cellStyleTable);


        //Row 10
        Row row10 = sheet.createRow(9);

        Cell cell91 = row10.createCell(0);
        cell91.setCellStyle(cellStyleTable);

        Cell cell92 = row10.createCell(1);
        cell92.setCellStyle(cellStyleTable);

        Cell cell14 = row10.createCell(3);
        cell14.setCellValue("Thành công");
        final int cellrange14 = sheet.addMergedRegion( new CellRangeAddress(9,9,3,4));
        cell14.setCellStyle(cellStyleTable);

        Cell cell15 = row10.createCell(5);
        cell15.setCellValue("Thoái trả");
        final int cellrange15 = sheet.addMergedRegion( new CellRangeAddress(9,9,5,6));
        cell15.setCellStyle(cellStyleTable);

        Cell cell16 = row10.createCell(7);
        cell16.setCellValue("Tổng tiền");
        final int cellrange16 = sheet.addMergedRegion( new CellRangeAddress(9,10,7,7));
        cell16.setCellStyle(cellStyleTable);

        Cell cell93 = row10.createCell(9);
        cell93.setCellStyle(cellStyleTable);

        //Row 11
        Row row11 = sheet.createRow(10);

        Cell cell94 = row11.createCell(1);
        cell94.setCellStyle(cellStyleTable);

        Cell cell95 = row11.createCell(2);
        cell95.setCellStyle(cellStyleTable);

        Cell cell96 = row11.createCell(7);
        cell96.setCellStyle(cellStyleTable);

        Cell cell97 = row11.createCell(8);
        cell97.setCellStyle(cellStyleTable);

        Cell cell98 = row11.createCell(9);
        cell98.setCellStyle(cellStyleTable);

        Cell cell17 = row11.createCell(3);
        cell17.setCellValue("SLGD");
        cell17.setCellStyle(cellStyleTable);

        Cell cell18 = row11.createCell(4);
        cell18.setCellValue("Số tiền");
        cell18.setCellStyle(cellStyleTable);

        Cell cell19 = row11.createCell(5);
        cell19.setCellValue("SLGD");
        cell19.setCellStyle(cellStyleTable);

        Cell cell20 = row11.createCell(6);
        cell20.setCellValue("Số tiền");
        cell20.setCellStyle(cellStyleTable);

        //Row 12
        Row row12 = sheet.createRow(11);
        Cell cell21 = row12.createCell(0);
        cell21.setCellStyle(cellStyleTable);

        Cell cell22 = row12.createCell(1);
        cell22.setCellStyle(cellStyleTable);

        Cell cell23 = row12.createCell(2);
        cell23.setCellValue("(1)");
        cell23.setCellStyle(cellStyleTable);

        Cell cell24 = row12.createCell(3);
        cell24.setCellStyle(cellStyleTable);

        Cell cell25 = row12.createCell(4);
        cell25.setCellValue("(2)");
        cell25.setCellStyle(cellStyleTable);

        Cell cell26 = row12.createCell(5);
        cell26.setCellStyle(cellStyleTable);

        Cell cell27 = row12.createCell(6);
        cell27.setCellValue("(3)");
        cell27.setCellStyle(cellStyleTable);

        Cell cell28 = row12.createCell(7);
        cell28.setCellValue("(4)=(2)-(3)");
        cell28.setCellStyle(cellStyleTable);

        Cell cell29 = row12.createCell(8);
        cell29.setCellValue("(5)");
        cell29.setCellStyle(cellStyleTable);

        Cell cell30= row12.createCell(9);
        cell30.setCellValue("(6)=(1)+(4)-(5)");
        cell30.setCellStyle(cellStyleTable);

        //Row 13
        Row row13 = sheet.createRow(12);
        row13.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell31 = row13.createCell(0);
        cell31.setCellValue("Gói chu kỳ ngắn");
        final int cellrange31 = sheet.addMergedRegion( new CellRangeAddress(12,13,0,0));
        cell31.setCellStyle(cellStyleTable);

        Cell cell32 = row13.createCell(1);
        cell32.setCellValue("Thuê bao phát triển mới");
        cell32.setCellStyle(cellStyleTable);

        Cell cell33 = row13.createCell(2);
        cell33.setCellStyle(cellStyleTable);

        Cell cell34 = row13.createCell(3);
        cell34.setCellStyle(cellStyleTable);

        Cell cell35 = row13.createCell(4);
        cell35.setCellStyle(cellStyleTable);

        Cell cell36 = row13.createCell(5);
        cell36.setCellStyle(cellStyleTable);

        Cell cell37 = row13.createCell(6);
        cell37.setCellStyle(cellStyleTable);

        Cell cell38 = row13.createCell(7);
        cell38.setCellStyle(cellStyleTable);

        Cell cell39 = row13.createCell(8);
        cell39.setCellStyle(cellStyleTable);

        Cell cell40 = row13.createCell(9);
        cell40.setCellStyle(cellStyleTable);

        //Row 14
        Row row14 = sheet.createRow(13);
        row14.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell41 = row14.createCell(1);
        cell41.setCellValue("Thuê bao hiện hữu");
        cell41.setCellStyle(cellStyleTable);

        Cell cell42 = row14.createCell(2);
        cell42.setCellStyle(cellStyleTable);

        Cell cell43 = row14.createCell(3);
        cell43.setCellStyle(cellStyleTable);

        Cell cell44 = row14.createCell(4);
        cell44.setCellStyle(cellStyleTable);

        Cell cell45 = row14.createCell(5);
        cell45.setCellStyle(cellStyleTable);

        Cell cell46 = row14.createCell(6);
        cell46.setCellStyle(cellStyleTable);

        Cell cell47 = row14.createCell(7);
        cell47.setCellStyle(cellStyleTable);

        Cell cell48 = row14.createCell(8);
        cell48.setCellStyle(cellStyleTable);

        Cell cell49 = row14.createCell(9);
        cell49.setCellStyle(cellStyleTable);

        //Row 15
        Row row15 = sheet.createRow(14);
        row15.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell50 = row15.createCell(0);
        cell50.setCellValue("Gói chu kỳ dài");
        cell50.setCellStyle(cellStyleTable);

        Cell cell51 = row15.createCell(1);
        cell51.setCellStyle(cellStyleTable);

        Cell cell52 = row15.createCell(2);
        cell52.setCellStyle(cellStyleTable);

        Cell cell53 = row15.createCell(3);
        cell53.setCellStyle(cellStyleTable);

        Cell cell54 = row15.createCell(4);
        cell54.setCellStyle(cellStyleTable);

        Cell cell55 = row15.createCell(5);
        cell55.setCellStyle(cellStyleTable);

        Cell cell56 = row15.createCell(6);
        cell56.setCellStyle(cellStyleTable);

        Cell cell57 = row15.createCell(7);
        cell57.setCellStyle(cellStyleTable);

        Cell cell58 = row15.createCell(8);
        cell58.setCellStyle(cellStyleTable);

        Cell cell59 = row15.createCell(9);
        cell59.setCellStyle(cellStyleTable);

        //Row 16
        Row row16 = sheet.createRow(15);
        row16.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell60 = row16.createCell(0);
        cell60.setCellValue("Tổng cộng");
        final int cellrange60 = sheet.addMergedRegion( new CellRangeAddress(15,15,0,1));
        cell60.setCellStyle(cellStyleTable);

        Cell cell99 = row16.createCell(1);
        cell99.setCellStyle(cellStyleTable);

        Cell cell61 = row16.createCell(2);
        cell61.setCellStyle(cellStyleTable);

        Cell cell62 = row16.createCell(3);
        cell62.setCellStyle(cellStyleTable);

        Cell cell63 = row16.createCell(4);
        cell63.setCellStyle(cellStyleTable);

        Cell cell64 = row16.createCell(5);
        cell64.setCellStyle(cellStyleTable);

        Cell cell65 = row16.createCell(6);
        cell65.setCellStyle(cellStyleTable);

        Cell cell66 = row16.createCell(7);
        cell66.setCellStyle(cellStyleTable);

        Cell cell67 = row16.createCell(8);
        cell67.setCellStyle(cellStyleTable);

        Cell cell68 = row16.createCell(9);
        cell68.setCellStyle(cellStyleTable);

        //Row 17
        Row row17 = sheet.createRow(16);
        Cell cell69 = row17.createCell(0);
        cell69.setCellValue("Số còn phải trả cho ….. T…../2021 (Bằng chữ):Một triệu, ba trăm năm mươi bảy nghìn đồng./.");
        final int cellrange69 = sheet.addMergedRegion( new CellRangeAddress(16,16,0,9));
        cell69.setCellStyle(cellStyleItalic);

        //Row 18
        Row row18 = sheet.createRow(17);
        Cell cell70 = row18.createCell(0);
        cell70.setCellValue("Biên bản này được lập thành 04 bản có giá trị pháp lý như nhau, VNPT Media giữ 02 bản, …......giữ 02 bản.");
        final int cellrange70 = sheet.addMergedRegion( new CellRangeAddress(17,17,0,9));
        cell70.setCellStyle(cellStyle7);

        //Row 19
        Row row19 = sheet.createRow(18);
        Cell cell71 = row19.createCell(6);
        cell71.setCellValue("Hà Nội, ngày 05 tháng 12 năm 2021");
        final int cellrange71 = sheet.addMergedRegion( new CellRangeAddress(18,18,6,9));
        cell71.setCellStyle(cellStyle8);

        //Row 20
        Row row20 = sheet.createRow(19);
        Cell cell72 = row20.createCell(1);
        cell72.setCellValue("ĐẠI DIỆN BÊN A");
        final int cellrange72 = sheet.addMergedRegion( new CellRangeAddress(19,19,1,3));
        cell72.setCellStyle(cellStyleValue);

        Cell cell73 = row20.createCell(6);
        cell73.setCellValue("ĐẠI DIỆN BÊN B");
        final int cellrange73 = sheet.addMergedRegion( new CellRangeAddress(19,19,6,9));
        cell73.setCellStyle(cellStyleValue);

        //Row 21
        Row row21 = sheet.createRow(20);
        row21.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell74 = row21.createCell(6);
        cell74.setCellValue(" BP. Đối Soát");
        cell74.setCellStyle(cellStyleValue);

        Cell cell75 = row21.createCell(7);
        cell75.setCellValue("Phòng Vận hành");
        cell75.setCellStyle(cellStyleValue);

        Cell cell76 = row21.createCell(8);
        cell76.setCellValue("TUQ TỔNG GIÁM ĐỐC\n" +
                "PGĐ VNPT FINTECH");
        final int cellrange76 = sheet.addMergedRegion( new CellRangeAddress(20,20,8,9));
        cell76.setCellStyle(cellStyleValue);


    }

    private void writeDataRow() {

    }

    public void exportExcelReportVNP(HttpServletResponse response) throws IOException {
        writeHeaderRow();
        writeDataRow();
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        outputStream.close();
    }


}
