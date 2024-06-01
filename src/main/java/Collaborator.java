import javax.mail.Message;
import java.util.List;

public class Collaborator {
    public static void main(String[] args) {
        ReceiveEmails receiveEmails = new ReceiveEmails();
        WriteToExcel write = new WriteToExcel();
        List<Message> emails;
        emails = receiveEmails.getMessages();
        write.createWorkBook(emails);

    }
}
