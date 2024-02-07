import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class GoogleSheetsExample {
    private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "token";
    //Connecting to the spreadsheet ID
    private static final String spreadsheetId = "1aKM2F_wa7q6QSFCsHNW9bCCmf8nFrk1EJya5sAhZNIU";

    private static final List<String> SCOPES =
            Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

    /**
     * Creates an authorized Credential object.
     *
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT)
            throws IOException {
        // Load client secrets.
        InputStream in = GoogleSheetsExample.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets =
                GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();

        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();

        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.

        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

        System.out.println("Defining the spreadsheet cell range");
        final String range = "engenharia_de_software!A4:H27";

        System.out.println("Building a new instance of Sheets considering the API from Google Workspace API");
        Sheets service =
                new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                        .setApplicationName(APPLICATION_NAME)
                        .build();

        System.out.println("Accessing the values from spreadsheet");
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();

        System.out.println("Getting the values accessed in the spreadsheet and adding them to a list");
        List<List<Object>> values = response.getValues();

        //cellNumber is a cell behind the first cell will be updated in the spreadsheet
        int cellNumber = 3;

        System.out.println("Verifying if the list of values is null...");
        if (values == null || values.isEmpty()) {
            System.out.println("No data found");
        }
        else {
            System.out.println("Looping through all rows in the list and calculating the student's average grade " +
                    "to verify their situations");
            for (List<Object> row : values) {
                int classes = 60;

                int absences = Integer.parseInt((String) row.get(2));
                double p1 = Double.parseDouble((String) row.get(3));
                double p2 = Double.parseDouble((String) row.get(4));
                double p3 = Double.parseDouble((String) row.get(5));
                double averageGrade = Math.round((p1 + p2 + p3) / 3);

                System.out.println("\n");
                System.out.println(row.get(1) +"'s average grade is " + averageGrade);

                int scoreToBeApproved = 0;
                cellNumber = cellNumber + 1;

                int percentageOfAbsences = (absences/ classes) * 100;

                if (percentageOfAbsences > 25) {
                    System.out.println("The situation of " + row.get(1) + " is 'Reprovado por Falta'. Percentage of " +
                            "absences: " + percentageOfAbsences);

                    updateValues("Reprovado por Falta", scoreToBeApproved, cellNumber, service);
                }
                else if (averageGrade < 50) {
                    System.out.println("The situation of " + row.get(1) + " is 'Reprovado por Nota'");

                    updateValues("Reprovado por Nota", scoreToBeApproved, cellNumber, service);
                }
                else if (averageGrade < 70) {
                    scoreToBeApproved = (int) ((50 * 2) - averageGrade);

                    System.out.println("The situation of " + row.get(1) + " is 'Exame Final'. It's necessary " +
                            scoreToBeApproved + " points in the final test for approval");

                    updateValues("Exame Final", scoreToBeApproved, cellNumber, service);
                }
                else {
                    System.out.println("The situation of " + row.get(1) + " is 'Aprovado'");

                    updateValues("Aprovado", scoreToBeApproved, cellNumber, service);
                }

            }
            System.out.println("\nUpdated values");
        }

    }

    private static void updateValues(String situation, int naf, int cellNumber, Sheets service)
            throws IOException {

        System.out.println("Setting the variables that will fill the cells");
        ValueRange body = new ValueRange()
                .setValues(List.of(
                        Arrays.asList(situation, naf)
                ));

        System.out.println("Accessing the specific cell to update the cells with the variables");

        service.spreadsheets().values()
                .update(spreadsheetId, "engenharia_de_software!G" + cellNumber + ":H" + cellNumber, body)
                .setValueInputOption("RAW")
                .execute();
    }
}