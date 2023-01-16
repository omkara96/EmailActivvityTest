import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import com.auxilii.msgparser.*;
import com.auxilii.msgparser.MsgParser;
import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import java.util.Iterator;
import java.util.List;

import static java.lang.System.exit;

public class MainActivity {

    private static String msg_file =null;
    private String out_path=null;

    public static void main(String[] args) {

        if (args.length == 0){
            System.out.println("Please Provide valid no of Arguments no arguments are specified !");
            System.out.println("1st required argument is Msg file");
            System.out.println("2nd required argument is Output Path");
            exit(1);
        } else if (args.length>2){
            System.out.println(" Error !!! Too many arguments Required 2 you provided " + args.length);
            exit(1);
        } else{
            System.out.println("Starting the activity");
            System.out.println("Argument Length: " + args.length);
            System.out.println("Initializing .MSG extraction class");
            MainActivity mainActivity = new MainActivity();
            System.out.println("Starting Reading Contents of .Msg File");
            mainActivity.Read_MsgFile(msg_file);
            try {
                System.out.println("Starting Extracting Attachments from .Msg file");
                mainActivity.Save_TestAttachment();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    private void Read_MsgFile(String msg_file){
        try {
            MsgParser msgp = new MsgParser();
            Message msg = msgp.parseMsg(msg_file);
            String from_email = msg.getFromEmail();
            String from_name = msg.getFromName();
            String subject = msg.getSubject();
            String body = msg.getBodyText();
            String to_list = msg.getDisplayTo();
            String cc_list = msg.getDisplayCc();
            String bcc_list = msg.getDisplayBcc();
            List list = msg.getAttachments();
            System.out.println("Attachments -" + list.size());
            Iterator it_list = list.iterator();
            Object attachemetn = null;
            while (it_list.hasNext()) {
                attachemetn = it_list.next();
                System.out.println(attachemetn);
            }
            System.out.println("-----");
            System.out.println("from_email " + from_email);
            System.out.println("from_name " + from_name);
            System.out.println("to_list " + to_list);
            System.out.println("cc_list " + cc_list);
            System.out.println("bcc_list " + bcc_list);
            System.out.println("subject " + subject);
            System.out.println("body " + body);
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    private int Save_TestAttachment() throws IOException {

        MAPIMessage message = new MAPIMessage("D:\\Test\\Fwd_book.msg");
        AttachmentChunks[] attachments = message.getAttachmentFiles();

        if (attachments.length > 0) {
           // File d = new File("D:\\Attachments");

                for (AttachmentChunks attachment : attachments) {
                    String file_name = attachment.attachFileName.getValue();
                    System.out.println("Attachment Name: " + file_name);
                    byte[] data = attachment.attachData.getValue();
                    attachment.getEmbeddedAttachmentObject().toString();
                    String path = "D:\\Attachments\\" + file_name;
                    File nfile = new File(path);
                    writeBytesToFile(path, data);

            }
                return 1;
        } else{
            return 0;
        }

    }

    private static void writeBytesToFile(String fileOutput, byte[] bytes)
            throws IOException {

        try (FileOutputStream fos = new FileOutputStream(fileOutput)) {
            fos.write(bytes);
        }

    }
}
