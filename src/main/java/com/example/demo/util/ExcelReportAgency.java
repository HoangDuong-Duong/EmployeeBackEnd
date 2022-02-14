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

public class ExcelReportAgency {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<Employee> employeeList;

    public ExcelReportAgency(List<Employee> employees) {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("ReportAgency");
        this.employeeList = employees;

    }

    private void writeHeaderRow() {

        CellStyle cellStyle = workbook.createCellStyle();
        XSSFFont fontCellHeader = workbook.createFont();
        fontCellHeader.setBold(true);
        fontCellHeader.setFontName("Times New Roman");
        fontCellHeader.setFontHeight(13);
        cellStyle.setFont(fontCellHeader);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle.setWrapText(true);


        CellStyle cellStyle10 = workbook.createCellStyle();
        XSSFFont fontCellHeader33 = workbook.createFont();
        fontCellHeader33.setBold(true);
        fontCellHeader33.setFontName("Times New Roman");
        fontCellHeader33.setFontHeight(16);
        cellStyle10.setFont(fontCellHeader33);
        cellStyle10.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle10.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle10.setWrapText(true);


        CellStyle cellStyle2 = workbook.createCellStyle();
        XSSFFont fontCellHeader2 = workbook.createFont();
        fontCellHeader2.setBold(true);
        fontCellHeader2.setFontName("Times New Roman");
        fontCellHeader2.setFontHeight(10);
        cellStyle2.setFont(fontCellHeader2);
        cellStyle2.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyle2.setWrapText(true);


        CellStyle cellStyle3 = workbook.createCellStyle();
        XSSFFont fontCellHeader3 = workbook.createFont();
        fontCellHeader3.setFontName("Times New Roman");
        fontCellHeader3.setFontHeight(12);
        fontCellHeader3.setItalic(true);
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


        CellStyle cellStyleItalicBold = workbook.createCellStyle();
        XSSFFont fontCellHeader11= workbook.createFont();
        fontCellHeader11.setFontName("Times New Roman");
        fontCellHeader11.setFontHeight(1);
        fontCellHeader11.setItalic(true);
        fontCellHeader11.setBold(true);
        cellStyleItalicBold.setFont(fontCellHeader11);
        cellStyleItalicBold.setAlignment(CellStyle.ALIGN_RIGHT);
        cellStyleItalicBold.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleItalicBold.setWrapText(true);


        CellStyle cellStyleValue = workbook.createCellStyle();
        XSSFFont fontCellHeader5 = workbook.createFont();
        fontCellHeader5.setFontName("Times New Roman");
        fontCellHeader5.setFontHeight(12);
        cellStyleValue.setFont(fontCellHeader5);
        cellStyleValue.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleValue.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleValue.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleValue.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleValue.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleValue.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyleValue.setWrapText(true);


        CellStyle cellStyleTable = workbook.createCellStyle();
        XSSFFont fontCellHeader7 = workbook.createFont();
        fontCellHeader7.setBold(true);
        fontCellHeader7.setFontName("Times New Roman");
        fontCellHeader7.setFontHeight(12);
        cellStyleTable.setFont(fontCellHeader7);
        cellStyleTable.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleTable.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleTable.setWrapText(true);
        cellStyleTable.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleTable.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleTable.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleTable.setBorderLeft(HSSFCellStyle.BORDER_THIN);



        CellStyle cellStyleVlue = workbook.createCellStyle();
        XSSFFont fontCellHeader12 = workbook.createFont();
        fontCellHeader12.setBold(true);
        fontCellHeader12.setFontName("Times New Roman");
        fontCellHeader12.setFontHeight(12);
        cellStyleVlue.setFont(fontCellHeader12);
        cellStyleVlue.setAlignment(CellStyle.ALIGN_LEFT);
        cellStyleVlue.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleVlue.setWrapText(true);
        cellStyleVlue.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleVlue.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleVlue.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleVlue.setBorderLeft(HSSFCellStyle.BORDER_THIN);




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
        fontCellHeader9.setBold(true);
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
        row1.setHeightInPoints((4 * sheet.getDefaultRowHeightInPoints()));
        Cell cell = row1.createCell(0);
        cell.setCellValue("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM"+"\n"+"Độc lập - Tự do - Hạnh phúc"+"\n"+"-------o0o--------");
        final int cellrange1 = sheet.addMergedRegion( new CellRangeAddress(0,0,0,7));
        cell.setCellStyle(cellStyle);

        //Row2
        Row row2 = sheet.createRow(1);
        row2.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell1 = row2.createCell(0);
        cell1.setCellValue("BIÊN BẢN TỔNG HỢP SỐ LIỆU VÀ PHÍ GIAO DỊCH DỊCH VỤ THU HỘ GÓI CƯỚC");
        final int cellrange2 = sheet.addMergedRegion( new CellRangeAddress(1,1,0,7));
        cell1.setCellStyle(cellStyle10);

        //Row3
        Row row3 = sheet.createRow(2);
        Cell cell2 = row3.createCell(0);
        cell2.setCellValue("GIỮA  TỔNG CÔNG TY TRUYỀN THÔNG VÀ Đối tác");
        final int cellrange3 = sheet.addMergedRegion( new CellRangeAddress(2,2,0,7));
        cell2.setCellStyle(cellStyle2);

        //Row 4
        Row row4 = sheet.createRow(3);
        Cell cell3 = row4.createCell(0);
        cell3.setCellValue("Từ: 00:00:00 ngày 01/07/2021 đến 23:59:59 ngày 31/07/2021");
        final int cellrange4 = sheet.addMergedRegion( new CellRangeAddress(3,3,0,7));
        cell3.setCellStyle(cellStyle3);

        //Row 5
        Row row5 = sheet.createRow(4);
        row5.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell4 = row5.createCell(0);
        cell4.setCellValue(" - Căn cứ theo hợp đồng cung cấp dịch vụ thu hộ cước trả sau giữa TỔNG CÔNG TY TRUYỀN THÔNG và Đối tác số hợp đồng 123 ");
        final int cellrange5 = sheet.addMergedRegion( new CellRangeAddress(4,4,0,7));
        cell4.setCellStyle(cellStyleItalic);

        //Row 6
        Row row6 = sheet.createRow(5);
        row6.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell5 = row6.createCell(0);
        cell5.setCellValue(" - Hôm nay, ngày     tháng      năm 2021, TỔNG CÔNG TY TRUYỀN THÔNG và Đối tác cùng ký xác nhận  " +
                "Biên bản tổng hợp số liệu và phí giao dịch dịch vụ thu hộ cước trả sau tháng 07/2021 cụ thể như sau:");
        final int cellrange6 = sheet.addMergedRegion( new CellRangeAddress(5,5,0,7));
        cell5.setCellStyle(cellStyleItalic);


        //Row 7
        Row row7 = sheet.createRow(6);
        Cell cell7 = row7.createCell(7);
        cell7.setCellValue("ĐVT: VNĐ");
        cell7.setCellStyle(cellStyleItalicBold);


        //Row 8
        Row row8 = sheet.createRow(7);
        Cell cell8 = row8.createCell(0);
        cell8.setCellValue("STT");
        cell8.setCellStyle(cellStyleTable);

        Cell cell9 = row8.createCell(1);
        cell9.setCellValue("Loại gói cước");
        cell9.setCellStyle(cellStyleTable);

        Cell cell10 = row8.createCell(2);
        cell10.setCellValue("Nội dung");
        cell10.setCellStyle(cellStyleTable);

        Cell cell11 = row8.createCell(3);
        cell11.setCellValue("Số lượng Giao dịch");
        cell11.setCellStyle(cellStyleTable);

        Cell cell12 = row8.createCell(4);
        cell12.setCellValue("Giá trị giao dịch");
        cell12.setCellStyle(cellStyleTable);

        Cell cell13 = row8.createCell(5);
        cell13.setCellValue("Mức phí");
        cell13.setCellStyle(cellStyleTable);

        Cell cell14 = row8.createCell(6);
        cell14.setCellValue("Số tiền phí\n" +
                "(Không bao gồm VAT)");
        cell14.setCellStyle(cellStyleTable);

        Cell cell15 = row8.createCell(7);
        cell15.setCellValue("Số tiền phí\n" +
                "(Đã bao gồm VAT)");
        cell15.setCellStyle(cellStyleTable);

        //Row 9
        Row row9 = sheet.createRow(8);
        Cell cell16 = row9.createCell(0);
        cell16.setCellValue("(1)");
        cell16.setCellStyle(cellStyleTable);

        Cell cell17 = row9.createCell(1);
        cell17.setCellValue("(2)");
        cell17.setCellStyle(cellStyleTable);

        Cell cell18 = row9.createCell(2);
        cell18.setCellStyle(cellStyleTable);

        Cell cell19 = row9.createCell(3);
        cell19.setCellValue("(3)");
        cell19.setCellStyle(cellStyleTable);

        Cell cell20 = row9.createCell(4);
        cell20.setCellValue("(4)");
        cell20.setCellStyle(cellStyleTable);

        Cell cell21 = row9.createCell(5);
        cell21.setCellValue("(5)");
        cell21.setCellStyle(cellStyleTable);

        Cell cell22 = row9.createCell(6);
        cell22.setCellValue("(6)=(4)*(5)/1.1");
        cell22.setCellStyle(cellStyleTable);

        Cell cell23 = row9.createCell(7);
        cell23.setCellValue("(7)=(4)*(5)");
        cell23.setCellStyle(cellStyleTable);

        //Row 10
        Row row10 = sheet.createRow(9);
        row10.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell24 = row10.createCell(0);
        cell24.setCellValue("1");
        cell24.setCellStyle(cellStyleValue);

        Cell cell25 = row10.createCell(1);
        cell25.setCellValue("Gói ngắn ngày");
        final int cellrange25 = sheet.addMergedRegion( new CellRangeAddress(9,10,1,1));
        cell25.setCellStyle(cellStyleValue);

        Cell cell26 = row10.createCell(2);
        cell26.setCellValue("Thuê bao mới đăng ký");
        cell26.setCellStyle(cellStyleVlue);

        Cell cell27 = row10.createCell(3);
        cell27.setCellStyle(cellStyleValue);

        Cell cell28 = row10.createCell(4);
        cell28.setCellStyle(cellStyleValue);

        Cell cell29 = row10.createCell(5);
        cell29.setCellStyle(cellStyleValue);

        Cell cell30 = row10.createCell(6);
        cell30.setCellStyle(cellStyleValue);

        Cell cell31 = row10.createCell(7);
        cell31.setCellStyle(cellStyleValue);

        //Row 11
        Row row11 = sheet.createRow(10);
        row11.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell32 = row11.createCell(0);
        cell32.setCellValue("2");
        cell32.setCellStyle(cellStyleValue);

        Cell cell33 = row11.createCell(1);
        cell33.setCellStyle(cellStyleValue);

        Cell cell34 = row11.createCell(2);
        cell34.setCellValue("Thuê bao hiện hữu đăng ký");
        cell34.setCellStyle(cellStyleVlue);

        Cell cell35 = row11.createCell(3);
        cell35.setCellStyle(cellStyleValue);

        Cell cell36 = row11.createCell(4);
        cell36.setCellStyle(cellStyleValue);

        Cell cell37 = row11.createCell(5);
        cell37.setCellStyle(cellStyleValue);

        Cell cell38 = row11.createCell(6);
        cell38.setCellStyle(cellStyleValue);

        Cell cell39 = row11.createCell(7);
        cell39.setCellStyle(cellStyleValue);

        //Row 12
        Row row12 = sheet.createRow(11);
        row12.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell40 = row12.createCell(0);
        cell40.setCellValue("3");
        cell40.setCellStyle(cellStyleValue);

        Cell cell41 = row12.createCell(1);
        cell41.setCellValue("Gói dài ngày");
        final int cellrange41 = sheet.addMergedRegion( new CellRangeAddress(11,12,1,1));
        cell41.setCellStyle(cellStyleValue);

        Cell cell42 = row12.createCell(2);
        cell42.setCellValue("Thuê bao mới đăng ký");
        cell42.setCellStyle(cellStyleVlue);

        Cell cell43 = row12.createCell(3);
        cell43.setCellStyle(cellStyleValue);

        Cell cell44 = row12.createCell(4);
        cell44.setCellStyle(cellStyleValue);

        Cell cell45 = row12.createCell(5);
        cell45.setCellStyle(cellStyleValue);

        Cell cell46 = row12.createCell(6);
        cell46.setCellStyle(cellStyleValue);

        Cell cell47 = row12.createCell(7);
        cell47.setCellStyle(cellStyleValue);


        //Row 13
        Row row13 = sheet.createRow(12);
        row13.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell48 = row13.createCell(0);
        cell48.setCellValue("4");
        cell48.setCellStyle(cellStyleValue);

        Cell cell49 = row13.createCell(1);
        cell49.setCellStyle(cellStyleValue);

        Cell cell50 = row13.createCell(2);
        cell50.setCellValue("Thuê bao hiện hữu đăng ký");
        cell50.setCellStyle(cellStyleVlue);

        Cell cell51 = row13.createCell(3);
        cell51.setCellStyle(cellStyleValue);

        Cell cell52 = row13.createCell(4);
        cell52.setCellStyle(cellStyleValue);

        Cell cell53 = row13.createCell(5);
        cell53.setCellStyle(cellStyleValue);

        Cell cell54 = row13.createCell(6);
        cell54.setCellStyle(cellStyleValue);

        Cell cell55 = row13.createCell(7);
        cell55.setCellStyle(cellStyleValue);


        //Row thiếu
        Row row100 = sheet.createRow(13);
        row100.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell100 = row100.createCell(0);
        final int cellrange58 = sheet.addMergedRegion( new CellRangeAddress(13,13,0,2));
        cell100.setCellValue("Tổng cộng");
        cell100.setCellStyle(cellStyleTable);

        Cell cell101 = row100.createCell(1);
        cell101.setCellStyle(cellStyleTable);

        Cell cell102 = row100.createCell(2);
        cell102.setCellStyle(cellStyleTable);

        Cell cell103 = row100.createCell(3);
        cell103.setCellStyle(cellStyleTable);

        Cell cell104 = row100.createCell(4);
        cell104.setCellStyle(cellStyleTable);

        Cell cell105 = row100.createCell(5);
        cell105.setCellStyle(cellStyleTable);

        Cell cell106 = row100.createCell(6);
        cell106.setCellStyle(cellStyleTable);

        Cell cell107 = row100.createCell(7);
        cell107.setCellStyle(cellStyleTable);


        //Row 14
        Row row14 = sheet.createRow(15);
        Cell cell56 = row14.createCell(0);
        final int cellrange56 = sheet.addMergedRegion( new CellRangeAddress(15 ,15,0,6));
        cell56.setCellValue("Tổng số tiền Đối tác cần thanh toán cho TỔNG CÔNG TY TRUYỀN THÔNG tháng 07/2021 là:");
        cell56.setCellStyle(cellStyle7);

        Cell cell57 = row14.createCell(7);
        cell57.setCellValue("");
        cell57.setCellStyle(cellStyle3);


        //Row 15
        Row row15 = sheet.createRow(16);
        Cell cell58 = row15.createCell(0);
        final int cellrange60 = sheet.addMergedRegion( new CellRangeAddress(16,16,0,7));
        cell58.setCellValue("Số tiền bằng chữ:");
        cell58.setCellStyle(cellStyleItalic);

        //Row 16
        Row row16 = sheet.createRow(17);
        Cell cell59 = row16.createCell(0);
        final int cellrange59 = sheet.addMergedRegion( new CellRangeAddress(17 ,17,0,6));
        cell59.setCellValue("Tổng số tiền phí dịch vụ thu hộ gói cước Đối tác được hưởng tháng 07/2021 là:");
        cell59.setCellStyle(cellStyle7);

        Cell cell60 = row16.createCell(7);
        cell60.setCellValue("");
        cell60.setCellStyle(cellStyle3);

        //Row 17
        Row row17 = sheet.createRow(18);
        Cell cell61 = row17.createCell(0);
        final int cellrange61 = sheet.addMergedRegion( new CellRangeAddress(18 ,18,0,7));
        cell61.setCellValue("Số tiền bằng chữ:");
        cell61.setCellStyle(cellStyleItalic);


        //Row 18
        Row row18 = sheet.createRow(20);
        Cell cell62 = row18.createCell(5);
        final int cellrange62 = sheet.addMergedRegion( new CellRangeAddress(20 ,20,5,7));
        cell62.setCellValue("Hà Nội, ngày    tháng     năm 2021");
        cell62.setCellStyle(cellStyle3);

        //Row 19
        Row row19 = sheet.createRow(21);
        row19.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        Cell cell63 = row19.createCell(0);
        final int cellrange63 = sheet.addMergedRegion( new CellRangeAddress(21,21,0,4));
        cell63.setCellValue("ĐẠI DIỆN VNPT-MEDIA");
        cell63.setCellStyle(cellStyle8);

        Cell cell64 = row19.createCell(5);
        final int cellrange64 = sheet.addMergedRegion( new CellRangeAddress(21 ,21,5,7));
        cell64.setCellValue("ĐẠI DIỆN ĐỐI TÁC");
        cell64.setCellStyle(cellStyle8);

        //Row 20
        Row row20 = sheet.createRow(22);
        Cell cell65 = row20.createCell(0);
        final int cellrange65 = sheet.addMergedRegion( new CellRangeAddress(22 ,22,0,1));
        cell65.setCellValue("PHÓ GIÁM ĐỐC VNPT \n FINTECH");
        cell65.setCellStyle(cellStyle8);

        Cell cell66 = row20.createCell(2);
        cell66.setCellValue("P.KĨ THUẬT");
        cell66.setCellStyle(cellStyle8);

        Cell cell67 = row20.createCell(3);
        final int cellrange67 = sheet.addMergedRegion( new CellRangeAddress(22 ,22,3,4));
        cell67.setCellValue("P.ĐỐI SOÁT & HTKH");
        cell67.setCellStyle(cellStyle8);

        //Row 21
        Row row21 = sheet.createRow(29);
        Cell cell68 = row21.createCell(0);
        final int cellrange68 = sheet.addMergedRegion( new CellRangeAddress(29 ,29,0,1));
        cell68.setCellValue("Trần Thị Mai Liên");
        cell68.setCellStyle(cellStyle8);

        Cell cell69 = row21.createCell(2);
        cell69.setCellValue("Nguyễn Ngọc Khánh");
        cell69.setCellStyle(cellStyle8);

        Cell cell70 = row21.createCell(3);
        final int cellrange70 = sheet.addMergedRegion( new CellRangeAddress(29 ,29,3,4));
        cell70.setCellValue("Trần Trung Nghĩa");
        cell70.setCellStyle(cellStyle8);


    }

    private void writeDataRow() {

    }

    public void exportExcelReportAgency(HttpServletResponse response) throws IOException {
        writeHeaderRow();
        writeDataRow();
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        outputStream.close();
    }


}
