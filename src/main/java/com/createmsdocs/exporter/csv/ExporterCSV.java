package com.createmsdocs.exporter.csv;

import com.createmsdocs.dto.MemberDTO;
import org.supercsv.io.CsvBeanWriter;
import org.supercsv.io.ICsvBeanWriter;
import org.supercsv.prefs.CsvPreference;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class ExporterCSV {

    private List<MemberDTO> list;

    public ExporterCSV(List<MemberDTO> list) {
        this.list = list;
    }

    public void export(HttpServletResponse response) throws IOException {
        setResponse(response);
        writeFile(response);
    }

    public static void setResponse(HttpServletResponse response){
        response.setContentType("text/csv");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=users_" + currentDateTime + ".csv";
        response.setHeader(headerKey, headerValue);
    }

    private void writeFile(HttpServletResponse response) throws IOException {
        ICsvBeanWriter csvWriter = new CsvBeanWriter(response.getWriter(), CsvPreference.STANDARD_PREFERENCE);
        String[] csvHeader = {"Member ID", "E-mail", "First Name", "Last Name"};
        String[] nameMapping = {"id", "email", "firstName", "lastName"};

        csvWriter.writeHeader(csvHeader);

        for (MemberDTO tempMember : list) {
            csvWriter.write(tempMember, nameMapping);
        }

        csvWriter.close();
    }
}
