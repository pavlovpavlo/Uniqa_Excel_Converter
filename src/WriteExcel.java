import org.apache.poi.hssf.OldExcelFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;

public class WriteExcel
{




    private static XSSFCellStyle createStyleForTitle(XSSFWorkbook workbook) {
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }
    public static void create_file(String name, File source) throws IOException {
        Workbook wb;
        Sheet sheet = null;
        int j=0;
        try
        {
            FileInputStream in = new FileInputStream("C:\\ExcelConvert\\List_Demo.xls");

            wb = new HSSFWorkbook(in);
            sheet = wb.getSheetAt(0);
        }
        catch (Exception e)
        {
           //e.printStackTrace();
            throw  new IOException("oops");
        }
        Iterator<Row> it = sheet.iterator();
        ArrayList<String> array = new ArrayList<>();

        while (it.hasNext())
        {

            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                if (cell.getCellType() != CellType.BLANK && row.getRowNum() != 0) {

                    if (sheet.getRow(0).getCell(cell.getColumnIndex()).getStringCellValue().indexOf("E-Mail") >= 0) {
                        //System.out.println(name +" "+(sheet.getRow(j).getCell(1).getCellType()));
                        if(name.equals(sheet.getRow(j).getCell(1).getStringCellValue()))
                        {
                            if (sheet.getRow(j).getCell(3).getCellType() == CellType.NUMERIC) {
                                array.add ("" + (int) sheet.getRow(j).getCell(3).getNumericCellValue());
                               System.out.println("" + (int) sheet.getRow(j).getCell(3).getNumericCellValue());
                            }
                            if (sheet.getRow(j).getCell(3).getCellType() == CellType.STRING) {
                            array.add(cell.getStringCellValue());
                               // System.out.println("(cell.getStringCellValue()");
                        }
                        }

                    }

                }
            }


            j++;
        }
       /* FileChannel  sourceChannel = new FileInputStream(source).getChannel();
        FileChannel destChannel = null;*/
        try {


            for (int n = 0; n < array.size(); n++) {
                File dest = new File("C:\\File\\"+array.get(n)+".xlsx");
                Files.copy(source.toPath(), dest.toPath());
                /*destChannel = new FileOutputStream(dest).getChannel();
                destChannel.transferFrom(sourceChannel, 0, sourceChannel.size());*/
            }

        }
        catch (Exception e)
        {
            //source.delete();
        }

        finally {

            /*sourceChannel.close();
            destChannel.close();*/

        }
        //source.delete();
    }
    public static void write(ArrayList<UniquaTemplate> excellist,String name) throws IOException
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Лист1");

        ArrayList<UniquaTemplate> list = excellist;

        int rownum = 0;
        Cell cell;
        Row row;
        XSSFCellStyle style = createStyleForTitle(workbook);

        row = sheet.createRow(rownum);
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Код Морион / Код услуги");
        cell.setCellStyle(style);
        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("Наименование позиции поставщика");
        cell.setCellStyle(style);
        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("Цена НАЧАЛЬНАЯ (грн.)");
        cell.setCellStyle(style);
        cell = row.createCell(3, CellType.STRING);
        cell.setCellValue("Скидка в %");
        cell.setCellStyle(style);
        cell = row.createCell(4, CellType.STRING);
        cell.setCellValue("Цена ДЛЯ UNIQA (грн.)");
        cell.setCellStyle(style);
        cell = row.createCell(5, CellType.STRING);
        cell.setCellValue("Размер НДС");
        cell.setCellStyle(style);
        cell = row.createCell(6, CellType.STRING);
        cell.setCellValue("Доступное количество");
        cell.setCellStyle(style);
        for (UniquaTemplate emp : list) {
            rownum++;
            row = sheet.createRow(rownum);
            cell = row.createCell(0, CellType.NUMERIC);
            cell.setCellValue(emp.getArticle());

            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue(emp.getName());

            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue(emp.getPrice());
            cell = row.createCell(3, CellType.STRING);
            cell.setCellValue(emp.getDiscount());

            cell = row.createCell(4, CellType.STRING);
            cell.setCellValue(emp.getUniqaprice());

            cell = row.createCell(5, CellType.STRING);
            cell.setCellValue(emp.getNdsflag());

            cell = row.createCell(6, CellType.STRING);
            cell.setCellValue(emp.getQuantity());

        }

        File fiile = new File("C:\\File\\"+name+".xlsx");

        FileOutputStream outFile = new FileOutputStream(fiile);
        workbook.write(outFile);
        String file_name=name.substring(1,name.indexOf(";")-1);
        //System.out.println(name.indexOf(";"));
        create_file(file_name,new File(fiile.getAbsolutePath()));
        fiile.delete();
      }
}
