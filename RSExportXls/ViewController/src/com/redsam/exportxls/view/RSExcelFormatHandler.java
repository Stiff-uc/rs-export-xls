package com.redsam.exportxls.view;

import java.io.IOException;

import java.text.SimpleDateFormat;

import java.util.Date;

import javax.servlet.ServletOutputStream;

import oracle.adf.share.ADFContext;
import oracle.adf.view.rich.export.ExportContext;
import oracle.adf.view.rich.export.FormatHandler;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class RSExcelFormatHandler extends FormatHandler {
    public RSExcelFormatHandler() {
        super();
    }

    Workbook xlsWorkbook;
    Sheet xlsSheet;
    Row currentRow;
    Cell currentCell;
    CellStyle xlsColumnHeaderStyle;
    CellStyle xlsTitleStyle;
    int currentRowIndx = 0;
    int currentColumnIndx = 0;
    int maxColumns = 0;

    Boolean isHeaderColumn;

    int xlsCurrentRowSpan = 0;
    int xlsCurrentColumnSpan = 0;

    private static final String EXCEL_FORMAT_TYPE = "RSExcelExport";
    private static final String EXCEL_CONTENT_TYPE_PREFIX = "application/vnd.ms-excel;charset=";


    private ServletOutputStream _writer = null;


    @Override
    public String getFormatType() {
        return EXCEL_FORMAT_TYPE;
    }

    @Override
    public String getContentType(ExportContext exportContext) {
        return EXCEL_CONTENT_TYPE_PREFIX + exportContext.getCharset();
    }

    @Override
    public Object setupExportTarget(ExportContext exportContext) throws IOException {
        prepareExcelWorkbook();
        ServletOutputStream writer = (ServletOutputStream) createExportTarget(exportContext);
        _writer = writer;
        return writer;
    }

    // do not override standart hook
    @Override
    protected Object createExportTarget(ExportContext exportContext) throws IOException {
        ServletOutputStream outputStream = exportContext.getServletResponse().getOutputStream();
        return outputStream;
    }

    @Override
    public void preContent(ExportContext exportContext) throws IOException {
        currentRowIndx = 0;
        currentColumnIndx = 0;
        maxColumns = 0;

        Row row = xlsSheet.createRow(currentRowIndx++);
        row.setHeight((short) (256));
        Cell cell = row.createCell(0);
        cell.setCellValue(exportContext.getTitle());
        cell.setCellStyle(xlsTitleStyle);
    }

    @Override
    public void startTableElement(ExportContext exportContext) throws IOException {
    }

    @Override
    public void startRowElement(ExportContext exportContext) throws IOException {
        currentRow = xlsSheet.createRow(currentRowIndx);
    }

    @Override
    public void startTableHeaderElement(ExportContext exportContext) throws IOException {
        isHeaderColumn = true;
        currentCell = currentRow.createCell(currentColumnIndx);
        currentCell.setCellStyle(xlsColumnHeaderStyle);
    }

    @Override
    public void writeColSpan(ExportContext exportContext, int i) throws IOException {
    }

    @Override
    public void writeRowSpan(ExportContext exportContext, int i) throws IOException {
    }

    @Override
    @SuppressWarnings("oracle.jdeveloper.java.insufficient-catch-block")
    public void writeText(ExportContext exportContext, Object object, String string) throws IOException {
        double d = 0;
        Boolean isDouble = false;
        try {
            d = Double.parseDouble(String.valueOf(object));
            isDouble = true;
        } catch (NumberFormatException nfe) {
            // not a date
        }
        if (object != null) {
            if (isDouble) {
                currentCell.setCellValue(d);
            } else {
                currentCell.setCellValue(String.valueOf(object));
            }
        }
    }

    @Override
    public void endTableHeaderElement(ExportContext exportContext) throws IOException {
        isHeaderColumn = false;
        currentColumnIndx++;
    }

    @Override
    public void startDataElement(ExportContext exportContext) throws IOException {
        currentCell = currentRow.createCell(currentColumnIndx);
    }

    @Override
    public void endDataElement(ExportContext exportContext) throws IOException {
        currentColumnIndx++;
    }

    @Override
    public void endRowElement(ExportContext exportContext) throws IOException {
        if (maxColumns < currentColumnIndx) {
            maxColumns = currentColumnIndx;
        }
        currentColumnIndx = 0;
        currentRowIndx++;
    }

    @Override
    public void endTableElement(ExportContext exportContext) throws IOException {
    }

    @Override
    public void postContent(ExportContext exportContext) throws IOException {

        xlsSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, maxColumns - 1));
        for (int colNum = 0; colNum < maxColumns; colNum++) {
            xlsSheet.autoSizeColumn(colNum);
        }

        addSheetComment();

        xlsWorkbook.write(_writer);

        //        super.postContent(exportContext);
    }


    @Override
    public void beginWrapDetailItem(ExportContext exportContext, Object object) throws IOException {
    }

    @Override
    public void endWrapDetailItem(ExportContext exportContext) throws IOException {
    }

    @Override
    public void startDetailStamp(ExportContext exportContext) throws IOException {
    }

    @Override
    public void endDetailStamp(ExportContext exportContext) throws IOException {
    }

    private void prepareExcelWorkbook() {
        xlsWorkbook = new HSSFWorkbook();
        Font headerFont = xlsWorkbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);


        xlsColumnHeaderStyle = xlsWorkbook.createCellStyle();
        xlsColumnHeaderStyle.setAlignment(CellStyle.ALIGN_CENTER);
        xlsColumnHeaderStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        xlsColumnHeaderStyle.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
        xlsColumnHeaderStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        xlsColumnHeaderStyle.setFont(headerFont);

        Font boldFont = xlsWorkbook.createFont();
        boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        xlsTitleStyle = xlsWorkbook.createCellStyle();
        xlsTitleStyle.setFont(boldFont);
        xlsTitleStyle.setAlignment(CellStyle.ALIGN_RIGHT);
        xlsTitleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        xlsSheet = xlsWorkbook.createSheet("Export");
    }

    private void addSheetComment() {
        String message =
            "Created by " + ADFContext.getCurrent().getSecurityContext().getUserName() + " " +
            new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(new Date());


        CreationHelper factory = xlsWorkbook.getCreationHelper();
        Drawing drawing = xlsSheet.createDrawingPatriarch();
        ClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(0);
        anchor.setCol2(0);
        anchor.setRow1(0);
        anchor.setRow2(0);

        // Create the comment and set the text+author
        Comment comment = drawing.createCellComment(anchor);
        RichTextString str = factory.createRichTextString(message);
        comment.setString(str);
        comment.setAuthor(ADFContext.getCurrent().getSecurityContext().getUserName());
    }

}
