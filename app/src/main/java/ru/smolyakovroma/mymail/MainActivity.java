package ru.smolyakovroma.mymail;

import android.os.AsyncTask;
import android.os.StrictMode;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Locale;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class MainActivity extends AppCompatActivity {

    EditText etMail;
    EditText etTitle;
    EditText etMessage;
    Button btnSend, btnRead;
    Session session = null;


    String rec;
    String subject;
    String textMessage;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        //для чтение почты получаем разрешение
        StrictMode.ThreadPolicy policy = new StrictMode.ThreadPolicy.Builder().permitAll().build();
        StrictMode.setThreadPolicy(policy);

        setContentView(R.layout.activity_main);

        etMail = (EditText) findViewById(R.id.mail);
        etTitle = (EditText) findViewById(R.id.title);
        etMessage = (EditText) findViewById(R.id.message);
        btnSend = (Button) findViewById(R.id.send);
        btnRead = (Button) findViewById(R.id.read);

        btnSend.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                 rec = etMail.getText().toString();
                 subject = etTitle.getText().toString();
                 textMessage = etMessage.getText().toString();

                Properties props = new Properties();
                props.put("mail.smtp.host", "smtp.mail.ru");
                props.put("mail.smtp.socketFactory.port", "465");
                props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
                props.put("mail.smtp.auth", "true");
                props.put("mail.smtp.port", "465");

                session = Session.getDefaultInstance(props, new Authenticator() {
                    @Override
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication("lug1c@mail.ru", "1234Zz");
                    }
                });

                RetreivFeedTask retreivFeedTask = new RetreivFeedTask();
                retreivFeedTask.execute();


            }
        });

        btnRead.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                try {
                    new MailReader(getApplicationContext());
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (MessagingException e) {
                    e.printStackTrace();
                }
            }
        });

    }

    class RetreivFeedTask extends AsyncTask<String, Void, String>{

        @Override
        protected String doInBackground(String... params) {
            try{

                Message message = new MimeMessage(session);
                message.setFrom(new InternetAddress("lug1c@mail.ru"));
                message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(rec));
                message.setSubject("changes");
//                message.setContent(textMessage, "text/html; charset=utf-8");
//                message.setSentDate(new Date());

                // creates message part
                MimeBodyPart messageBodyPart1 = new MimeBodyPart();
                messageBodyPart1.setText("Message body");

                MimeBodyPart messageBodyPart3 = new MimeBodyPart();
                File file = new File(getFilesDir().toString(), "otchet.xls");
                getXLS(file);
                DataSource source2 = new FileDataSource(file.getPath());
                messageBodyPart3.setDataHandler(new DataHandler(source2));
                messageBodyPart3.setFileName(file.getPath());

                // creates multi-part
                Multipart multipart = new MimeMultipart();
                multipart.addBodyPart(messageBodyPart1);
                multipart.addBodyPart(messageBodyPart3);
                message.setContent(multipart);

                // sets the multi-part as e-mail's content
//                message.setContent(multipart, "text/plain");

                Transport.send(message);


            } catch (MessagingException e){
                e.printStackTrace();
            }

            return null;
        }

        @Override
        protected void onPostExecute(String aVoid) {
            etMail.setText("");
            etMessage.setText("");
            etTitle.setText("");
            Toast.makeText(getApplicationContext(), "eMail sent", Toast.LENGTH_LONG).show();
        }
    }

    public void getXLS(File file){
        WorkbookSettings wbSettings = new WorkbookSettings();
        wbSettings.setLocale(new Locale("en", "EN"));
        WritableWorkbook workbook;

        try {
            workbook = Workbook.createWorkbook(file, wbSettings);
            //Excel sheet name. 0 represents first sheet
            WritableSheet sheet = workbook.createSheet("MyShoppingList", 0);

            try {
                sheet.addCell(new Label(0, 0, "Subject")); // column and row
                sheet.addCell(new Label(1, 0, "Description"));

                for (int i = 1; i < 11; i++) {
                    sheet.addCell(new Label(0, i, "title "+i));
                    sheet.addCell(new Label(1, i, "desc "+i));
                }


            } catch (RowsExceededException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            }
            workbook.write();
            try {
                workbook.close();
            } catch (WriteException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
