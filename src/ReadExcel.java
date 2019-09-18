import org.apache.poi.hssf.OldExcelFormatException;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Optional;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;


public class ReadExcel {
    public static int countArt;
    public static int countPrice;
    public static int countMerge=0;
    public static ArrayList<UniquaTemplate> read(File name, String type) throws IOException, OldExcelFormatException {

        UniquaTemplate uniquaTemplate = new UniquaTemplate();
        ArrayList<UniquaTemplate> excellist= new ArrayList<UniquaTemplate>();
        Workbook wb;
        Sheet sheet = null;
        try
        {
            FileInputStream in = new FileInputStream(name);
            switch(type){
                case "xls":{
                    wb = new HSSFWorkbook(in);
                    sheet = wb.getSheetAt(0);
                }
                case "xlsx":{
                    wb = new XSSFWorkbook(in);
                    sheet =(XSSFSheet) wb.getSheetAt(0);
                }
                case "zip":{

                }
                case "csv":{

                }
            }
        }
        catch (Exception e) {
            throw new IOException("oops");
        }
        Iterator<Row> it = sheet.iterator();
        while (it.hasNext()) {
            Row row = it.next();
            countMerge=0;
            Iterator<Cell> cells = row.iterator();


            int firstrow=0;
            excellist.add(checkCells(row,sheet,cells,uniquaTemplate));

        }
        return excellist;
    }
    protected static UniquaTemplate checkCells(Row row, Sheet sheet, Iterator<Cell> cells, UniquaTemplate uniquaTemplate){


        while (cells.hasNext())
        {

            Cell cell = cells.next();
            for (int i = 0; i < sheet.getNumMergedRegions(); i++)
            {
                CellRangeAddress region = sheet.getMergedRegion(i);
                int rowNum = region.getFirstRow();
                if (rowNum == cell.getRowIndex() )
                {
                    countMerge=rowNum;
                }
            }
                /*Cell cellf = row.getCell(cell.getColumnIndex(), Row.RETURN_NULL_AND_BLANK);
                System.out.println(cell.);*/
                /*if ((cellf == null) || (cellf.equals("")) || (cellf.getCellType() == cellf.CELL_TYPE_BLANK))
                {
                    firstrow++;
                    System.out.println(firstrow);
                    break;
                }*/
               /* else*/
        }
            return uniquaTemplate;
    }
    public static Optional<Integer> tryParseInteger(String string) {
        try {
            return Optional.of(Integer.valueOf(string));
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }
}