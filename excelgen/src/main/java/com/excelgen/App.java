package com.excelgen;

/**
 * JXLS Excel Generation Demo
 */
public final class App {
    private App() {
    }

    /**
     * Demonstrates JXLS template processing with conditional logic.
     * @param args The arguments of the program.
     */
    public static void main(String[] args) {
        try {
            String templatePath = "person_template.xlsx";

            // Create the template
            TemplateCreator.createTemplate(templatePath);

            // Test Case 1: Person under 18 (should show parent name)
            Person minor = new Person("John Doe", 15, "Jane Doe");
            String outputMinor = "person_minor_output.xlsx";
            TemplateCreator.processTemplate(templatePath, outputMinor, minor);
            System.out.println("Generated Excel for minor: " + outputMinor);

            // Test Case 2: Person over 18 (should NOT show parent name)
            Person adult = new Person("Alice Smith", 25, "Bob Smith");
            String outputAdult = "person_adult_output.xlsx";
            TemplateCreator.processTemplate(templatePath, outputAdult, adult);
            System.out.println("Generated Excel for adult: " + outputAdult);

            System.out.println("\nExcel generation completed successfully!");

        } catch (Exception e) {
            System.err.println("Error processing template: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
