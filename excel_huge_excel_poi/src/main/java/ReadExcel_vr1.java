import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.concurrent.TimeUnit;
import java.util.stream.Stream;
import org.apache.commons.lang3.time.StopWatch;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;

/**
 * To file
 */

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;

public class ReadExcel_vr1 {

    public static void main(String[] args) throws IOException {

        try (InputStream is = new FileInputStream("/jcortessFiles/coding/huge_excel_files/500000_Records_Data_Test.xlsx");
             ReadableWorkbook wb = new ReadableWorkbook(is)) {

            StopWatch watch = new StopWatch();
            watch.start();
            wb.getSheets().forEach(sheet ->
            {
                try (Stream<Row> rows = sheet.openStream()) {

                    final String jsonOutputTest = "{";

                    /**
                     * Using rows.skip(1).parallel().forEach()--> Processing time :: 11426 milliseconds
                     */
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
                        String path = "/jcortessFiles/coding/huge_excel_files/500000_Records_Data_Test.json";

                        try {
                            File file = new File(path);
                            boolean fileExists = file.exists();

                            FileWriter fw = new FileWriter(file, true); // true means to append to the file
                            BufferedWriter bw = new BufferedWriter(fw);

                            if (fileExists) {
                                bw.newLine(); // add a new line before appending the content
                            }
                            bw.write(jsonOutputTest2);
                            bw.close();
                            fw.close();
                            //System.out.println("String appended to file successfully!");
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                        //System.out.println(jsonOutputTest2);
                        /**
                         * Using concat --> Processing time :: 4501 milliseconds
                         */
                        /*AtomicReference<String> jsonOutputTest2 = new AtomicReference<>("");
                        AtomicReference<String> jsonOutputTestFinal = new AtomicReference<>("");
                        jsonOutputTest2.set(jsonOutputTest.concat(
                                "\"Name\"" + ":" + "\"" + name + "\"" + "," +
                                        "\"Country\"" + ":" + "\"" + country + "\"" + "," +
                                        "\"item\"" + ":" + "\"" + item + "\"" + "," +
                                        "\"sales\"" + ":" + "\"" + sales + "\"" + "," +
                                        "\"order\"" + ":" + "\"" + order + "\"" + "," +
                                        "\"orderDate\"" + ":" + "\"" + orderDate + "\"" + "," +
                                        "\"orderId\"" + ":" + orderId + "," +
                                        "\"ship\"" + ":" + "\"" + ship + "\"" + "," +
                                        "\"unit\"" + ":" + unit + "," +
                                        "\"unitP\"" + ":" + unitP + "," +
                                        "\"unitC\"" + ":" + unitC + "," +
                                        "\"totalR\"" + ":" + totalR + "," +
                                        "\"totalC\"" + ":" + totalC + "," +
                                        "\"totalP\"" + ":" + totalP +
                                        "}"));
                        jsonOutputTestFinal.set(jsonOutputTest2.get().concat("]"));
                        System.out.println(jsonOutputTestFinal);*/
                        /**
                         * Using system.out--> Processing time :: 9646 milliseconds
                         */
                        /*System.out.println("Cell str value :: " + name);
                        System.out.println("Cell str value :: " + country);
                        System.out.println("Cell str value :: " + item);
                        System.out.println("Cell str value :: " + sales);
                        System.out.println("Cell str value :: " + order);
                        System.out.println("Cell str value :: " + orderDate);
                        System.out.println("Cell str value :: " + orderId);
                        System.out.println("Cell str value :: " + ship);
                        System.out.println("Cell str value :: " + unit);
                        System.out.println("Cell str value :: " + unitP);
                        System.out.println("Cell str value :: " + unitC);
                        System.out.println("Cell str value :: " + totalR);
                        System.out.println("Cell str value :: " + totalC);
                        System.out.println("Cell str value :: " + totalP);*/

                    });


                } catch (Exception e) {
                    e.printStackTrace();
                }
                watch.stop();
                System.out.println("Processing time :: " + watch.getTime(TimeUnit.MILLISECONDS));
            });
        }
    }
}