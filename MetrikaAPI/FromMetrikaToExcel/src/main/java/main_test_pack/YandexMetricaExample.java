package main_test_pack;
import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class YandexMetricaExample {
    // Replace these values with your own
    private static final String token = "AQAAAAA...";
    private static final String counterId = "12345678";
    private static final String period = "2021-01-01%3A2021-01-31";

    // The URL for sending requests to the Yandex Metrica API
    private static final String apiUrl = "https://api-metrika.yandex.net/stat/v1/data";

    // The name of the Excel file to save the results
    private static final String fileName = "yandex_metrica_data.xlsx";

    public static void main(String[] args) {
        try {
            // Create a new HTTP connection
            HttpURLConnection connection = (HttpURLConnection) new URL(apiUrl).openConnection();

            // Set the request method to POST
            connection.setRequestMethod("POST");

            // Set the request headers
            connection.setRequestProperty("Authorization", "OAuth " + token);
            connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
            connection.setDoOutput(true);

            // Write the request body with the query parameters
            DataOutputStream out = new DataOutputStream(connection.getOutputStream());
            out.writeBytes("ids=" + counterId + "&metrics=ym:s:visits&dimensions=ym:s:lastsignTrafficSource&date1=" + period + "&date2=" + period);
            out.flush();
            out.close();

            // Get the response code and message
            int responseCode = connection.getResponseCode();
            String responseMessage = connection.getResponseMessage();

            // Check if the response is OK (200)
            if (responseCode == HttpURLConnection.HTTP_OK) {
                System.out.println("Request successful: " + responseMessage);

                // Read the response body as a string
                BufferedReader in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
                StringBuilder responseBody = new StringBuilder();
                String inputLine;

                while ((inputLine = in.readLine()) != null) {
                    responseBody.append(inputLine);
                }
                in.close();

                System.out.println("Response body: " + responseBody.toString());

                // Parse the response body as a JSON object
                JsonObject jsonObject = JsonParser.parseString(responseBody.toString()).getAsJsonObject();

                // Get the data array from the JSON object
                JsonArray dataArray = jsonObject.getAsJsonArray("data");

                // Create a new Excel workbook and sheet
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet("Yandex Metrica Data");

                int rowNum = 0; // The current row number

                // Iterate over each element of the data array
                for (JsonElement element : dataArray) {
                    JsonObject dataObject = element.getAsJsonObject(); // The data object for each element

                    JsonArray dimensionsArray = dataObject.getAsJsonArray("dimensions"); // The dimensions array for each data object

                    JsonObject dimensionObject =
                            dimensionsArray.get(0).getAsJsonObject(); // The dimension object for each dimensions array (only one in this case)

                    String dimensionName =
                            dimensionObject.get("name").getAsString(); // The name of the dimension (the traffic source)

                    JsonArray metricsArray =
                            dataObject.getAsJsonArray("metrics"); // The metrics array for each data object

                    double metricValue =
                            metricsArray.get(0).getAsDouble(); // The value of the metric (the number of visits)

                    // Create a new row in the Excel sheet
                    Row row = sheet.createRow(rowNum++);

                    // Create two cells in the row: one for the dimension name and one for the metric value
                    Cell cell1 = row.createCell(0);
                    cell1.setCellValue(dimensionName);

                    Cell cell2 = row.createCell(1);
                    cell2.setCellValue(metricValue);
                }

                // Write the workbook to a file
                FileOutputStream outStream = new FileOutputStream(new File(fileName));
                workbook.write(outStream);
                outStream.close();

                System.out.println("Excel file created successfully.");

            } else {
                System.out.println("Request failed: " + responseMessage);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
/*
Итоговый excel файл сохраняется в ту же папку, где находится ваш проект.
Вы можете указать другой путь для сохранения файла, изменив строку кода:
FileOutputStream outStream = new FileOutputStream(new File(fileName));
*/