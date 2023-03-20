import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.concurrent.TimeUnit;
import java.util.stream.Stream;
import org.apache.commons.lang3.time.StopWatch;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import com.fasterxml.jackson.databind.node.ObjectNode;






/**
 * To file
 */

public class ReadExcelV2 {

    public static void main(String[] args) throws IOException {

        try (InputStream is = new FileInputStream("/jcortessFiles/coding/huge_excel_files/500000_Records_Data_Test.xlsx");
             ReadableWorkbook wb = new ReadableWorkbook(is)) {

            String path = "/jcortessFiles/coding/huge_excel_files/500000_Records_Data_Test.json";
            File file = new File(path);
            boolean fileExists = file.exists();

            //FileWriter fw = new FileWriter(file, true); // true means to append to the file
            //BufferedWriter bw = new BufferedWriter(fw);
            /**
             * BufferedWriter can be wrapped in a PrintWriter for better performance.
             */
            BufferedWriter bw = new BufferedWriter(new FileWriter(path));
            PrintWriter pw = new PrintWriter(bw);


            StopWatch watch = new StopWatch();
            watch.start();
            wb.getSheets().forEach(sheet ->
            {
                try (Stream<Row> rows = sheet.openStream()) {

                    final String jsonOutputTest = "{";
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

                        /**
                         * Using StringBuilder --> Processing time :: 4489 milliseconds
                         */
                        StringBuilder sb = new StringBuilder(jsonOutputTest);
                        sb.append("\"Name\"").append(":").append("\"").append(name).append("\"").append(",")
                                .append("\"Country\"").append(":").append("\"").append(country).append("\"").append(",")
                                .append("\"item\"").append(":").append("\"").append(item).append("\"").append(",")
                                .append("\"sales\"").append(":").append("\"").append(sales).append("\"").append(",")
                                .append("\"order\"").append(":").append("\"").append(order).append("\"").append(",")
                                .append("\"orderDate\"").append(":").append("\"").append(orderDate).append("\"").append(",")
                                .append("\"orderId\"").append(":").append(orderId).append(",")
                                .append("\"ship\"").append(":").append("\"").append(ship).append("\"").append(",")
                                .append("\"unit\"").append(":").append(unit).append(",")
                                .append("\"unitP\"").append(":").append(unitP).append(",")
                                .append("\"unitC\"").append(":").append(unitC).append(",")
                                .append("\"totalR\"").append(":").append(totalR).append(",")
                                .append("\"totalC\"").append(":").append(totalC).append(",")
                                .append("\"totalP\"").append(":").append(totalP)
                                .append("}");
                        String jsonOutputTest2;
                        String result = sb.toString();
                        jsonOutputTest2 = result;


                        /**
                         * With  System.out.println("String appended to file successfully!"); Processing time :: 12026
                         * Without  System.out.println("String appended to file successfully!"); Processing time :: 9563
                         */
                        try {
                            if (fileExists) {
                                bw.newLine(); // add a new line before appending the content
                            }
                            pw.write(jsonOutputTest2);


                            //System.out.println("String appended to file successfully!");
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    });
                    //Close the buffer writer and filewriter at the end for better performance
                    bw.close();
                    pw.close();
                    //fw.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
                watch.stop();
                System.out.println("Processing time :: " + watch.getTime(TimeUnit.MILLISECONDS));
            });
        }
    }
}