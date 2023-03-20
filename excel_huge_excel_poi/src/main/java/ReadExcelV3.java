import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.stream.Stream;

public class ReadExcelV3 {
    public static void main(String[] args) throws IOException {

        ObjectMapper objectMapper = new ObjectMapper();

        try (InputStream is = new FileInputStream("/jcortessFiles/coding/huge_excel_files/500000_Records_Data_Test.xlsx");
             ReadableWorkbook wb = new ReadableWorkbook(is)) {

            String path = "/jcortessFiles/coding/huge_excel_files/500000_Records_Data_Test.json";
            File file = new File(path);
            boolean fileExists = file.exists();

            FileOutputStream fileOutputStream = new FileOutputStream(file, true);

            wb.getSheets().forEach(sheet -> {

                try (Stream<Row> rows = sheet.openStream()) {

                    rows.skip(1).forEach(r -> {
                        String name = r.getCellAsString(0).orElse(null);
                        String country = r.getCellAsString(1).orElse(null);
                        String item = r.getCellAsString(2).orElse(null);
                        String sales = r.getCellAsString(3).orElse(null);
                        String order = r.getCellAsString(4).orElse(null);
                        LocalDateTime orderDate = r.getCellAsDate(5).orElse(null);
                        BigDecimal orderId = r.getCellAsNumber(6).orElse(null);
                        LocalDateTime ship = r.getCellAsDate(7).orElse(null);
                        BigDecimal unit = r.getCellAsNumber(8).orElse(null);
                        BigDecimal unitP = r.getCellAsNumber(9).orElse(null);
                        BigDecimal unitC = r.getCellAsNumber(10).orElse(null);
                        BigDecimal totalR = r.getCellAsNumber(11).orElse(null);
                        BigDecimal totalC = r.getCellAsNumber(12).orElse(null);
                        BigDecimal totalP = r.getCellAsNumber(13).orElse(null);

                        ObjectNode jsonObject = objectMapper.createObjectNode();

                        jsonObject.put("Name", name);
                        jsonObject.put("Country", country);
                        jsonObject.put("item", item);
                        jsonObject.put("sales", sales);
                        jsonObject.put("order", order);
                        jsonObject.put("orderDate", orderDate.toString());
                        jsonObject.put("orderId", orderId.doubleValue());
                        jsonObject.put("ship", ship.toString());
                        jsonObject.put("unit", unit.doubleValue());
                        jsonObject.put("unitP", unitP.doubleValue());
                        jsonObject.put("unitC", unitC.doubleValue());
                        jsonObject.put("totalR", totalR.doubleValue());
                        jsonObject.put("totalC", totalC.doubleValue());
                        jsonObject.put("totalP", totalP.doubleValue());

                        try {
                            if (fileExists) {
                                fileOutputStream.write(",".getBytes()); // add a comma before appending the content
                            }
                            objectMapper.writeValue(fileOutputStream, jsonObject);
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    });
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });

            //Close the fileoutputstream at the end for better performance
            fileOutputStream.close();
        }
    }
}
