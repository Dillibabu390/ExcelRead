import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.ArrayList;
public class Excel {
    public static void main(String[] args) {
        getRowcount();
    }
    private static void getRowcount() {
try {
    String excelpath = "C:\\Users\\Dillibabu\\Documents\\list.xlsx";
    XSSFWorkbook workbook = new XSSFWorkbook(excelpath);
    XSSFSheet sheet = workbook.getSheet("Sheet1");
    int rowcount = sheet.getPhysicalNumberOfRows();
    ArrayList<String> output = new ArrayList<>();
    for (int i =6 ;i< rowcount ; i++) {
        Row row = sheet.getRow(i);
        for (int j = 1; j < row.getLastCellNum(); j++) {
            String value = row.getCell(j).getStringCellValue();
            output.add(value);
        }
    }
    System.out.println(output);
}catch (Exception e){
    e.getMessage();
}
    }
}
