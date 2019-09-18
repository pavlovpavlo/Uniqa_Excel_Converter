import org.apache.poi.hssf.OldExcelFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Optional;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;


public class FileReader {

    public static UniquaTemplate uniquaTemplate;
    public static Workbook wb;
    public static Sheet sheet = null;
    public static int countPrice = 0;

    public static ArrayList<UniquaTemplate> read(File name, String type) throws IOException, OldExcelFormatException {
        ArrayList<UniquaTemplate> excellist= new ArrayList<UniquaTemplate>();
        int firstRow = 0;
        int firstCell = 0;
        try {
            FileInputStream in = new FileInputStream(name);
            switch(type){
                case "xls":{
                    wb = new HSSFWorkbook(in);
                    sheet = wb.getSheetAt(0);
                    break;
                }
                case "xlsx":{
                    wb = new XSSFWorkbook(in);
                    sheet =(XSSFSheet) wb.getSheetAt(0);
                    break;
                }
                case "zip":{
                    ZipOutput(name);
                    return null;
                }
                case "csv":{
                    try {
                        csvToXLSX(name.getName().substring(0,name.getName().lastIndexOf(".")), name);
                        FileInputStream incsv = new FileInputStream("C:\\FileAsConvert\\"+name.getName().substring(0,name.getName().lastIndexOf("."))+".xlsx");
                        wb = new XSSFWorkbook(incsv);
                        sheet =(XSSFSheet) wb.getSheetAt(0);
                    } catch (IOException e) {

                    }
                    break;
                }
            }

            for(int i = firstRow; i< sheet.getLastRowNum(); i++){
                for(int j = firstCell; j< sheet.getRow(i).getLastCellNum(); j++)
                {
                    // Код Мориона
                    if( sheet.getRow(i).getCell(j).getCellType() != CellType.BLANK && sheet.getRow(i).getRowNum() != 0 ) {
                        if (sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().indexOf("орион") > 0
                         || sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("код") >= 0) {
                            if (sheet.getRow(i).getCell(j).getCellType() == CellType.NUMERIC) {
                                uniquaTemplate.setArticle((int)sheet.getRow(i).getCell(j).getNumericCellValue());
                            }
                            if (sheet.getRow(i).getCell(j).getCellType() == CellType.STRING) {
                                try {
                                    uniquaTemplate.setArticle(tryParseInteger(""+sheet.getRow(i).getCell(j).getStringCellValue()).orElse(0));
                                }
                                catch (NumberFormatException e) {}
                            }
                        }
                        // Наименование
                        if (sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("вание") > 0
                           || sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().indexOf("овар") > 0
                           && sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("код") <0
                           || sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("азва") > 0) {
                            if(uniquaTemplate.getName().equals("")) {
                                if (sheet.getRow(i).getCell(j).getCellType() == CellType.NUMERIC) {
                                    uniquaTemplate.setName("" + sheet.getRow(i).getCell(j).getNumericCellValue());
                                }
                                if (sheet.getRow(i).getCell(j).getCellType() == CellType.STRING) {
                                    uniquaTemplate.setName(sheet.getRow(i).getCell(j).getStringCellValue());
                                }
                            }
                        }
                        // Цена начальная
                        if (sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("ена") > 0 && countPrice < 1
                          || sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("ціна") >= 0 && countPrice < 1) {
                            countPrice++;
                            if (sheet.getRow(i).getCell(j).getCellType() == CellType.NUMERIC) {
                                uniquaTemplate.setPrice("" + sheet.getRow(i).getCell(j).getNumericCellValue());
                                uniquaTemplate.setUniqaprice("" + sheet.getRow(i).getCell(j).getNumericCellValue());
                            }
                            if (sheet.getRow(i).getCell(j).getCellType() == CellType.STRING) {
                                uniquaTemplate.setPrice(sheet.getRow(i).getCell(j).getStringCellValue());
                                uniquaTemplate.setUniqaprice(sheet.getRow(i).getCell(j).getStringCellValue());
                            }
                        }
                        // Скидка
                        if (sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("идка") > 0
                         || sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("%") == 0) {
                            if (sheet.getRow(i).getCell(j).getCellType() == CellType.NUMERIC) {
                                uniquaTemplate.setDiscount(""+(int)((int)sheet.getRow(i).getCell(j).getNumericCellValue()<=0? sheet.getRow(i).getCell(j).getNumericCellValue()*100:sheet.getRow(i).getCell(j).getNumericCellValue())+"%");
                            }
                            if (sheet.getRow(i).getCell(j).getCellType() == CellType.STRING) {
                                uniquaTemplate.setDiscount(sheet.getRow(i).getCell(j).getStringCellValue());
                            }
                        }
                        // Цена для Уники
                        if (sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("ена") > 0 && countPrice > 0) {
                            countPrice++;
                            if(sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("ена с") >0 && sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("ена д")<1 ) {
                                if (sheet.getRow(i).getCell(j).getCellType() == CellType.NUMERIC) {
                                    uniquaTemplate.setPrice("" + sheet.getRow(i).getCell(j).getNumericCellValue());
                                }
                                if (sheet.getRow(i).getCell(j).getCellType() == CellType.STRING) {
                                    uniquaTemplate.setPrice(sheet.getRow(i).getCell(j).getStringCellValue());
                                }
                            }
                            else {
                                if (sheet.getRow(i).getCell(j).getCellType() == CellType.NUMERIC) {
                                    uniquaTemplate.setUniqaprice("" + sheet.getRow(i).getCell(j).getNumericCellValue());
                                }
                                if (sheet.getRow(i).getCell(j).getCellType() == CellType.STRING) {
                                    uniquaTemplate.setUniqaprice(sheet.getRow(i).getCell(j).getStringCellValue());
                                }
                            }
                        }
                        // НДС
                        if ((sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().indexOf("НДС") >= 0
                          || sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().indexOf("ПДВ") >= 0
                          || sheet.getRow(0).getCell(sheet.getRow(i).getCell(j).getColumnIndex()).getStringCellValue().toLowerCase().indexOf("авка") > 0)) {
                            if (sheet.getRow(i).getCell(j).getCellType() == CellType.NUMERIC) {
                                uniquaTemplate.setNdsflag(""+sheet.getRow(i).getCell(j).getNumericCellValue());
                            }
                            if (sheet.getRow(i).getCell(j).getCellType() == CellType.STRING) {
                                uniquaTemplate.setNdsflag(sheet.getRow(i).getCell(j).getStringCellValue());
                            }
                        }
                    }
                }
                excellist.add(uniquaTemplate);
            }
        }
        catch (Exception e) {
            throw new IOException("oops");
        }

        return excellist;
    }
    public static Optional<Integer> tryParseInteger(String string) {
        try {
            return Optional.of(Integer.valueOf(string));
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }
    protected static void ZipOutput(File names){
        Charset CP866 = Charset.forName("CP866");
        try(ZipInputStream zin = new ZipInputStream(new FileInputStream(names),CP866))
        {
            ZipEntry entry;
            String name;

            while((entry=zin.getNextEntry())!=null){
                name = entry.getName(); // получим название файла
                // распаковка
                FileOutputStream fout = new FileOutputStream("C:\\FileAsConvert\\" + name);
                for (int c = zin.read(); c != -1; c = zin.read()) {
                    fout.write(c);
                }
                File fiile = new File(name);
                WriteExcel.write(ReadExcel.read(fiile,fiile.getPath().substring(fiile.getPath().lastIndexOf(".")+1)),fiile.getName().substring(0,fiile.getName().lastIndexOf(".")));
                fiile.delete();
                fout.flush();
                zin.closeEntry();
                fout.close();
            }
        }
        catch(Exception ex){

            ex.printStackTrace();
        }
    }
    protected static void csvToXLSX(String name, File file)throws IOException {
        try {
            String csvFileAddress = file.getAbsolutePath(); //csv file address
            String xlsxFileAddress = "C:\\FileAsConvert\\"+name+".xlsx"; //xlsx file address

            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet("sheet1");
            String currentLine=null;
            int RowNum=0;
            BufferedReader br = new BufferedReader(new java.io.FileReader(csvFileAddress));
            while ((currentLine = br.readLine()) != null) {
                String str[] = currentLine.split(",");
                RowNum++;
                XSSFRow currentRow=sheet.createRow(RowNum);
                for(int i=0;i<str.length;i++){
                    currentRow.createCell(i).setCellValue(str[i]);
                }
            }

            FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Done");
        } catch (Exception ex) {
            System.out.println(ex.getMessage()+"Exception in try");
        }
    }
}