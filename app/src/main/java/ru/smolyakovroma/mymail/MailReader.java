package ru.smolyakovroma.mymail;

import android.content.Context;
import android.util.Log;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;
import javax.mail.search.SearchTerm;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class MailReader {
    Folder inbox;

    //Constructor of the calss.
    public MailReader(Context context) throws IOException, MessagingException {
 /*  Set the mail properties  */
        Properties props = System.getProperties();
        props.setProperty("mail.store.protocol", "imaps");

 /*  Create the session and get the store for read the mail. */
        Session session = Session.getDefaultInstance(props, null);
        Store store = session.getStore("imaps");
        store.connect("imap.mail.ru", "lug1c@mail.ru", "1234Zz");


        inbox = store.getFolder("Inbox");
        System.out.println("No of Unread Messages : " + inbox.getUnreadMessageCount());


        inbox.open(Folder.READ_ONLY);


        SearchTerm term = new SearchTerm() {
            @Override
            public boolean match(Message msg) {
                try {
                    if (msg.getSubject().contains("changes")) {
                        return true;
                    }
                } catch (MessagingException ex) {
                    ex.printStackTrace();
                }
                return false;
            }
        };

        List<File> attachments = new ArrayList<File>();
        Message messages[] = inbox.search(term);
        for (Message message : messages) {
            Multipart multipart = null;

            multipart = (Multipart) message.getContent();

            // System.out.println(multipart.getCount());

            for (int i = 0; i < multipart.getCount(); i++) {
                MimeBodyPart bodyPart = (MimeBodyPart) multipart.getBodyPart(i);
                if (bodyPart.getDisposition() != null && Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition())) {
                    InputStream is = bodyPart.getInputStream();
//                    File f = new File(context.getFilesDir().toString() + "/tmp.xls");

//                    Log.d("mail1", f.getPath().toString());


                    Workbook w;
                    try {
                        w = Workbook.getWorkbook(is);
                        // Get the first sheet
                        Sheet sheet = w.getSheet(0);
                        // Loop over first 10 column and lines

                            for (int k = 0; k < sheet.getRows(); k++) {
                                for (int j = 0; j < sheet.getColumns(); j++) {
                                Cell cell = sheet.getCell(j, k);
                                CellType type = cell.getType();
                                if (type == CellType.LABEL) {
                                    System.out.println("I got a label "
                                            + cell.getContents());
                                }

                                if (type == CellType.NUMBER) {
                                    System.out.println("I got a number "
                                            + cell.getContents());
                                }

                            }
                        }
                    } catch (BiffException e) {
                        e.printStackTrace();
                    }

//                    FileOutputStream fos = new FileOutputStream(f);
//                    byte[] buf = new byte[4096];
//                    int bytesRead;
//                    while ((bytesRead = is.read(buf)) != -1) {
//                        fos.write(buf, 0, bytesRead);
//                    }
//                    fos.close();
//                    attachments.add(f);
                }
            }

            inbox.close(true);
            store.close();


        }
    }
}