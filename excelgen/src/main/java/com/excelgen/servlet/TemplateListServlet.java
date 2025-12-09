package com.excelgen.servlet;

import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.jar.JarEntry;
import java.util.jar.JarFile;

public class TemplateListServlet extends HttpServlet {

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        response.setContentType("application/json");
        response.setCharacterEncoding("UTF-8");

        List<String> templates = getTemplatesFromResources();

        PrintWriter out = response.getWriter();
        out.print("[");
        for (int i = 0; i < templates.size(); i++) {
            out.print("\"" + templates.get(i) + "\"");
            if (i < templates.size() - 1) {
                out.print(",");
            }
        }
        out.print("]");
        out.flush();
    }

    private List<String> getTemplatesFromResources() {
        List<String> templates = new ArrayList<>();

        // Known templates - check which ones exist
        String[] knownTemplates = {
            "person_template.xlsx",
            "person_template_address.xlsx",
            "person_template_old_style.xlsx",
            "person_template_with_phones.xlsx"
        };

        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        if (classLoader == null) {
            classLoader = getClass().getClassLoader();
        }

        for (String template : knownTemplates) {
            try {
                InputStream is = classLoader.getResourceAsStream(template);
                if (is != null) {
                    templates.add(template);
                    is.close();
                }
            } catch (IOException e) {
                System.err.println("Error checking template: " + template);
                e.printStackTrace();
            }
        }

        return templates;
    }
}
