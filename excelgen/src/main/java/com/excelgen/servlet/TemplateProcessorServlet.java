package com.excelgen.servlet;

import com.excelgen.Address;
import com.excelgen.Person;
import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.MultipartConfig;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

@MultipartConfig
public class TemplateProcessorServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        try {
            // Get form parameters
            String name = request.getParameter("name");
            String ageStr = request.getParameter("age");
            String parentName = request.getParameter("parentName");
            String addressType = request.getParameter("addressType");
            String addressLine = request.getParameter("addressLine");
            String templateName = request.getParameter("template");

            // Debug logging
            System.out.println("Received parameters:");
            System.out.println("  name: " + name);
            System.out.println("  age: " + ageStr);
            System.out.println("  template: " + templateName);
            System.out.println("  Content-Type: " + request.getContentType());

            // Validate required parameters
            if (name == null || name.trim().isEmpty() ||
                ageStr == null || ageStr.trim().isEmpty() ||
                templateName == null || templateName.trim().isEmpty()) {

                response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
                response.getWriter().write("Missing required parameters: name=" + name + ", age=" + ageStr + ", template=" + templateName);
                return;
            }

            // Parse age
            int age;
            try {
                age = Integer.parseInt(ageStr);
            } catch (NumberFormatException e) {
                response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
                response.getWriter().write("Invalid age format");
                return;
            }

            // Create Person object
            Person person = new Person();
            person.setName(name);
            person.setAge(age);

            // Set parent name if provided (and not empty)
            if (parentName != null && !parentName.trim().isEmpty()) {
                person.setParentName(parentName);
            }

            // Create and set Address if provided
            if ((addressType != null && !addressType.trim().isEmpty()) ||
                (addressLine != null && !addressLine.trim().isEmpty())) {

                Address address = new Address();
                address.setType(addressType);
                address.setAddressLine(addressLine);
                person.setAddress(address);
            }

            // Load template from resources
            InputStream templateStream = getClass().getClassLoader()
                    .getResourceAsStream(templateName);

            if (templateStream == null) {
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                response.getWriter().write("Template not found: " + templateName);
                return;
            }

            // Set response headers for Excel download
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + "PersonReport_" + System.currentTimeMillis() + ".xls");

            // Get output stream
            OutputStream out = response.getOutputStream();

            // Create JXLS context and add person object
            Context context = new Context();
            context.putVar("person", person);

            // Process template using JXLS 2.x
            JxlsHelper.getInstance().processTemplate(templateStream, out, context);

            // Flush and close streams
            out.flush();
            templateStream.close();

        } catch (Exception e) {
            e.printStackTrace();
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().write("Error processing template: " + e.getMessage());
        }
    }

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        response.setStatus(HttpServletResponse.SC_METHOD_NOT_ALLOWED);
        response.getWriter().write("GET method not supported. Please use POST.");
    }
}
