import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.Store;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class ReceiveEmails {
    private static List<Message> emails;
    public List<Message> getMessages(){
        emails = new ArrayList<>();
        String host = "imap.gmail.com";
        String port = "993";
        String username = "work.ersunil69@gmail.com";
        String password = "";

        Properties properties = new Properties();
        properties.put("mail.imap.host", host);
        properties.put("mail.imap.port", port);
        properties.put("mail.imap.starttls.enable", "true");
        properties.put("mail.imap.ssl.trust", host);
        properties.put("mail.imap.ssl.enable", "true");

        Session emailSession = Session.getDefaultInstance(properties);

        try {
            // Create Imap store and connect to server
            Store store = emailSession.getStore("imap");
            store.connect(host, username, password);

            Folder inbox = store.getFolder("INBOX");
            inbox.open(Folder.READ_ONLY);

            // receiving messages from inbox
            Message[] messages = inbox.getMessages();
            for (int i = 0; i < messages.length; i++) {
                Message message = messages[i];
                emails.add(message);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return emails;
    }

}
