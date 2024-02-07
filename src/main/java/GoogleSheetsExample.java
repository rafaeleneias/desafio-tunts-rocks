import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

public class GoogleSheetsExample {
    private static final String SPREADSHEET_ID = "1aKM2F_wa7q6QSFCsHNW9bCCmf8nFrk1EJya5sAhZNIU";

    public static void main(String[] args) throws IOException, GeneralSecurityException {
        Sheets sheetsService = GoogleSheetsService.getSheetsService();

        // Ler informações da planilha
        List<List<Object>> values = readDataFromSheet(sheetsService, "Sheet1", "A1:B2");

        // Fazer cálculos
        double result = calculate(values);

        // Escrever resultado de volta na planilha
        writeResultToSheet(sheetsService, "Sheet1", "C1", result);
    }

    private static List<List<Object>> readDataFromSheet(Sheets service, String sheetName, String range) throws IOException {
        ValueRange response = service.spreadsheets().values()
                .get(SPREADSHEET_ID, sheetName + "!" + range)
                .execute();
        return response.getValues();
    }

    private static double calculate(List<List<Object>> values) {
        // Implemente seus cálculos aqui
        return 42.0; // Exemplo: retornando 42 como resultado
    }

    private static void writeResultToSheet(Sheets service, String sheetName, String range, double result) throws IOException {
        ValueRange body = new ValueRange().setValues(Collections.singletonList(Collections.singletonList(result)));
        service.spreadsheets().values()
                .update(SPREADSHEET_ID, sheetName + "!" + range, body)
                .setValueInputOption("RAW")
                .execute();
    }
}
