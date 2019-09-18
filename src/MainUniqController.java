import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import javax.swing.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

public class MainUniqController {
    static File fiile;
    static PrintWriter fileReader;
    public static List<File> fileList;
    public static void main(String[] args)
    {
        try
        {
            fileList= Files.walk(Paths.get("C:\\FileAsConvert")).filter(Files::isRegularFile).map(Path::toFile).collect(Collectors.toList());
            for(File fiile:fileList) {
                try {
                    WriteExcel.write(FileReader.read(fiile,fiile.getPath().substring(fiile.getPath().lastIndexOf(".")+1)),fiile.getName().substring(0,fiile.getName().lastIndexOf(".")));
                    fiile.delete();
                }
                catch (NullPointerException ex){

                }
                catch (Exception e) {
                    e.printStackTrace();
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        JOptionPane.showMessageDialog(null,"Operation Complite","Information",JOptionPane.INFORMATION_MESSAGE);

        System.exit(0);
    }
    public static void logExcel(String name, String log_text){
        Workbook wb = null;
        Sheet sheets = null;
        try
        {
            FileInputStream in = new FileInputStream("C:\\ExcelConvert\\List_Demo.xls");


                wb = new HSSFWorkbook(in);
                sheets = wb.getSheetAt(0);

        } catch (FileNotFoundException e) {
        } catch (IOException e) {

        }
        Iterator<Row> it = sheets.iterator();
        ArrayList<String> array = new ArrayList<>();

        while (it.hasNext())
        {
             Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                if (cell.getCellType() != CellType.BLANK && row.getRowNum() != 0) {
                    if (sheets.getRow(row.getRowNum()).getCell(6).getStringCellValue().indexOf(new SimpleDateFormat("dd.MM.yyyy").format(Calendar.getInstance().getTime()).toString().substring(0,10))<0)
                    {
                        System.out.println("trrtsss");
                        sheets.getRow(row.getRowNum()).getCell(6).setCellValue("");

                    }
                    if(name.toLowerCase().indexOf(sheets.getRow(row.getRowNum()).getCell(1).getStringCellValue().toLowerCase()) >0){
                        System.out.println("trrt");
                        sheets.getRow(row.getRowNum()).getCell(7).setCellValue(log_text);
                        sheets.getRow(row.getRowNum()).getCell(6).setCellValue(new SimpleDateFormat("dd.MM.yyyy").format(Calendar.getInstance().getTime()).toString().substring(0,10));
                        break;
                    }
                }
            }
        }
        try {
            FileOutputStream fileOut = new FileOutputStream("C:\\ExcelConvert\\List_Demo.xls");
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {

        }

    }
}

