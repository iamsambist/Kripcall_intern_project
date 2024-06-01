import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMultipart;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WriteToExcel {

    // workbook --> Sheet --> row --> Col

    public void createWorkBook(List<Message> emails){
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("ALl Emails");

        Row headerRow = sheet.createRow(0);
        String[] columns = {"Email Number", "From", "Subject","Body"};
        for(int i=0; i<4;i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
        }
        int rowNum = 1;
        for ( Message email : emails) {
            Row row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue(rowNum++);
            try {
                row.createCell(1).setCellValue(convertAddressesToString(email.getFrom()));
                row.createCell(2).setCellValue(email.getSubject());
                row.createCell(3).setCellValue(getTextFromMessage(email));
            } catch (Exception e) {
                throw new RuntimeException(e);
            }

        }

        String filepath = ".\\KripCall.xlsx";
        try {
            FileOutputStream outputStream = new FileOutputStream(filepath);
            workbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    private static String convertAddressesToString(Address[] addresses) {
        if (addresses == null || addresses.length == 0) {
            return "";
        }
        StringBuilder addressString = new StringBuilder();
        for (Address address : addresses) {
            if (address instanceof InternetAddress) {
                addressString.append(((InternetAddress) address).toUnicodeString()).append(", ");
            }
        }
        if (addressString.length() > 2) {
            addressString.setLength(addressString.length() - 2); // Remove trailing comma and space
        }
        return addressString.toString();
    }
    public static String getTextFromMessage(Message message) throws MessagingException, IOException {
        if (message.isMimeType("text/plain")) {
            return message.getContent().toString();
        } else if (message.isMimeType("multipart/*")) {
            MimeMultipart mimeMultipart = (MimeMultipart) message.getContent();
            return getTextFromMimeMultipart(mimeMultipart);
        }
        return "";
    }
    private static String getTextFromMimeMultipart(MimeMultipart mimeMultipart) throws MessagingException, IOException {
        StringBuilder result = new StringBuilder();
        int count = mimeMultipart.getCount();
        for (int i = 0; i < count; i++) {
            BodyPart bodyPart = mimeMultipart.getBodyPart(i);
            if (bodyPart.isMimeType("text/plain")) {
                result.append(bodyPart.getContent());
            } else if (bodyPart.isMimeType("text/html")) {
                String html = (String) bodyPart.getContent();
                result.append(org.jsoup.Jsoup.parse(html).text());
            } else if (bodyPart.getContent() instanceof MimeMultipart) {
                result.append(getTextFromMimeMultipart((MimeMultipart) bodyPart.getContent()));
            }
        }
        return result.toString();
    }
}
