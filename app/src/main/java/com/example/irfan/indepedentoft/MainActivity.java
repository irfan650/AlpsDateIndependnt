package com.example.irfan.indepedentoft;

import android.app.Activity;
import android.os.AsyncTask;
import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.Menu;
import android.view.MenuItem;

import com.independentsoft.exchange.Appointment;
import com.independentsoft.exchange.AppointmentPropertyPath;
import com.independentsoft.exchange.Attendee;
import com.independentsoft.exchange.Body;
import com.independentsoft.exchange.CalendarView;
import com.independentsoft.exchange.FindItemResponse;
import com.independentsoft.exchange.Folder;
import com.independentsoft.exchange.InstanceType;
import com.independentsoft.exchange.ItemId;
import com.independentsoft.exchange.RecurringMasterItemId;
import com.independentsoft.exchange.Service;
import com.independentsoft.exchange.ServiceException;
import com.independentsoft.exchange.StandardFolder;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;


public class MainActivity extends Activity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        //setContentView(R.layout.activity_my);

        new MyAsyncTask().execute();
    }


    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        //Inflate the menu; this adds items to the action bar if it is present.
        // getMenuInflater().inflate(R.menu.my, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        // Handle action bar item clicks here. The action bar will
        // automatically handle clicks on the Home/Up button, so long
        // as you specify a parent activity in AndroidManifest.xml.
        int id = item.getItemId();
        //if (id == R.id.action_settings) {
        return true;
    }
    //return super.onOptionsItemSelected(item);
}

class MyAsyncTask extends AsyncTask<String, Integer, String> {


    @Override
    protected String doInBackground(String... params) {
        // TODO Auto-generated method stub
        String s=postData(params);
        return s;
    }

    protected void onPostExecute(String result){

    }
    protected void onProgressUpdate(Integer... progress){

    }

    public String postData(String valueIWantToSend[]) {

        String returnValue = "";
        try {


            Service service = new Service("https://outlook.office365.com/EWS/Exchange.asmx", "RoomAdmin@scheduledisplay.com", "Meeting1234");
            System.out.println("connting ");
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            Date startTime = dateFormat.parse("2017-09-15 12:10:00");
            Date endTime = dateFormat.parse("2017-09-15 12:20:00");

            Date NewstartTime = dateFormat.parse("2017-09-13 16:00:00");
            Date NewendTime = dateFormat.parse("2017-09-13 16:30:00");


            // CreateEvent(startTime,endTime,"TEST1",service);
            // CreateEvent(startTime,endTime,"TEST5",service);
            // DeleteEvent(startTime,endTime,service);
            /// Meeting
            CreateMeeting(startTime,endTime,"update1",service);
            //UpdateMeeting(NewstartTime,NewendTime,"update1",service );
            // CancelMeeting("update1",service );
            //AcceptMeeting("Test Meeting",service );


            Date gstartTime = dateFormat.parse("2017-09-13 00:00:00");
            Date gendTime = dateFormat.parse("2017-09-17 00:00:00");
            GetEvents(gstartTime,gendTime, service);


            Folder inboxFolder = service.getFolder(StandardFolder.INBOX);

            Log.w("inboxFolder", inboxFolder.getDisplayName());

            returnValue = inboxFolder.getDisplayName();

        }
        catch (ServiceException ex)
        {
            Log.w("ServiceException", ":" + ex.getFaultCode());
            Log.w("ServiceException", ":" + ex.getFaultString());
            Log.w("ServiceException", ":" + ex.getMessage());
            Log.w("ServiceException", ":" + ex.getXmlMessage());
            Log.w("ServiceException", ":" + ex.getResponseCode());
            Log.w("ServiceException", ":" + ex.getRequestBody());
        }
        catch (Exception ex)
        {
            Log.w("Exception", ex.getMessage());
        }

        return returnValue;
    }

    private static void CreateMeeting(  Date startTime,   Date endTime, String Subject, Service service ) throws ServiceException {
        /// ADDD APPOINTMENT

        Appointment appointment = new Appointment();
        appointment.setSubject(Subject);
        appointment.setBody(new Body("Body text."));
        appointment.setStartTime(startTime);
        appointment.setEndTime(endTime);
        appointment.setLocation("room1");
        appointment.setReminderIsSet(true);
        appointment.setReminderMinutesBeforeStart(30);
        appointment.getRequiredAttendees().add(new Attendee("meetingroom@scheduledisplay.com"));
        //appointment.getOptionalAttendees().add(new Attendee("Mark@mydomain.com"));

        ItemId itemId = service.sendMeetingRequest(appointment);

    }

    private static void GetEvents(Date Startime,   Date Endtime, Service service ) throws ServiceException {
        int appnt_count = 0;


        try {
//            Service service = new Service("https://outlook.office365.com/EWS/Exchange.asmx", "RoomAdmin@scheduledisplay.com", "Meeting1234");
//
//            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.US);
//
//            Date startTime = dateFormat.parse("2017-09-10 16:00:00");
//            Date endTime = dateFormat.parse("2017-09-15 16:30:00");

            CalendarView view = new CalendarView(Startime, Endtime);
            System.out.println("Iam inside Exchange");
            FindItemResponse exchangeresponse = service.findItem(StandardFolder.CALENDAR, AppointmentPropertyPath.getAllPropertyPaths(), view);



            for (int i = 0; i < exchangeresponse.getItems().size(); i++) {
                Appointment appointment = (Appointment) exchangeresponse.getItems().get(i);

                System.out.println("Subject = " + appointment.getSubject());
                System.out.println("StartTime = " + appointment.getStartTime());
                System.out.println("EndTime = " + appointment.getEndTime());
                System.out.println("Body Preview = " + appointment.getBodyPlainText());
                System.out.println("----------------------------------------------------------------");
                appnt_count++;
                if (appointment.getInstanceType() == InstanceType.OCCURRENCE)
                {
                    RecurringMasterItemId masterId = new RecurringMasterItemId(appointment.getItemId().getId(), appointment.getItemId().getChangeKey());
                }
            }


        }catch (ServiceException e) {
            e.printStackTrace();

        }

        System.out.println("TOTAL =" + appnt_count);
    }

}




