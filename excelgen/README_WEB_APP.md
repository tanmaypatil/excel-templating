# Excel Template Generator - Web Application Guide

This guide explains how to run the Excel Template Generator as a web application using Apache Tomcat.

## Overview

The web application provides a user-friendly interface to:
- Select a person's name and age
- Optionally add parent name (for minors)
- Optionally add address information
- Select a JXLS template from available templates
- Download a processed Excel file with the data

## Prerequisites

Before running the web application, ensure you have:

1. **Java 17 or higher**
   ```bash
   java -version
   # Should show version 17 or higher
   ```

2. **Maven 3.6.3 or higher**
   ```bash
   mvn -version
   # Should show Maven 3.6.3 or higher
   ```

## Project Structure

```
excelgen/
├── src/
│   ├── main/
│   │   ├── java/com/excelgen/
│   │   │   ├── Person.java                         # Domain model
│   │   │   ├── Address.java                        # Address model
│   │   │   └── servlet/
│   │   │       ├── TemplateListServlet.java        # Lists available templates
│   │   │       └── TemplateProcessorServlet.java   # Processes templates
│   │   ├── resources/
│   │   │   ├── person_template.xlsx                # Basic template
│   │   │   ├── person_template_address.xlsx        # Template with address
│   │   │   ├── person_template_old_style.xlsx      # Legacy template
│   │   │   └── person_template_with_phones.xlsx    # Template with phones
│   │   └── webapp/
│   │       ├── index.html                          # Web form UI
│   │       └── WEB-INF/
│   │           └── web.xml                         # Servlet configuration
├── pom.xml                                          # Maven configuration
└── README_WEB_APP.md                                # This file
```

## Quick Start

### 1. Navigate to the Project Directory

```bash
cd excelgen
```

### 2. Build the Project

Build the WAR file (skipping tests for faster build):

```bash
mvn clean package -DskipTests
```

This will create `target/excelgen-1.0-SNAPSHOT.war`

### 3. Start the Tomcat Server

Run the embedded Tomcat 10 server:

```bash
mvn cargo:run
```

You should see output like:
```
[INFO] Starting Servlet engine: [Apache Tomcat/10.1.13]
[INFO] Starting ProtocolHandler ["http-nio-8080"]
```

### 4. Access the Application

Open your web browser and navigate to:

```
http://localhost:8080/excelgen/
```

## Using the Web Application

### Filling Out the Form

1. **Name** (required): Enter the person's name
   - Example: `John Doe`

2. **Age** (required): Enter the person's age
   - Example: `25`

3. **Parent Name** (optional): Enter parent's name if person is a minor
   - Example: `Jane Doe`
   - Note: This field is especially relevant for persons under 18

4. **Address Type** (optional): Select from dropdown
   - Options: Home, Office, Mailing, Other

5. **Address Line** (optional): Enter the address
   - Example: `123 Main Street, City, State 12345`

6. **Phone Numbers** (optional): Add multiple phone numbers
   - Click "+ Add Another Phone" to add more phone numbers
   - Each phone entry has:
     - Phone Type: Mobile, Home, Office, Other
     - Phone Number: e.g., `+1-555-1234`
   - Click "- Remove Last Phone" to remove the last phone entry
   - Multiple phone numbers will appear in templates that support them

7. **Select Template** (required): Choose from available templates
   - `person_template.xlsx` - Basic template with name, age, and conditional parent
   - `person_template_address.xlsx` - Includes address fields
   - `person_template_old_style.xlsx` - Legacy format
   - `person_template_with_phones.xlsx` - Includes phone numbers (supports multiple phones)

### Download the Report

Click the **"Download Excel Report"** button. The processed Excel file will be downloaded to your default downloads folder.

The filename will be in the format: `PersonReport_<timestamp>.xls`

## Template Processing Features

### Conditional Logic

Templates use JXLS 2.x conditional expressions:

- **Minor Check**: If age < 18, the parent name field is displayed
- **Adult Check**: If age >= 18, the parent name field is hidden
- **Address Display**: Address fields only appear if address data is provided

### Example Scenarios

**Scenario 1: Minor without Address**
```
Name: John Doe
Age: 15
Parent Name: Jane Doe
Template: person_template.xlsx
```
Output includes: Name, Age, Parent Name

**Scenario 2: Adult with Address**
```
Name: John Smith
Age: 30
Address Type: Home
Address Line: 456 Oak Avenue
Template: person_template_address.xlsx
```
Output includes: Name, Age, Address (Parent Name hidden)

**Scenario 3: Person with Multiple Phone Numbers**
```
Name: Sarah Johnson
Age: 28
Phone 1: Mobile - +1-555-1234
Phone 2: Office - +1-555-5678
Phone 3: Home - +1-555-9012
Template: person_template_with_phones.xlsx
```
Output includes: Name, Age, and all 3 phone numbers with their types

## Stopping the Server

To stop the Tomcat server:

1. Press `Ctrl+C` in the terminal where the server is running
2. Wait for the server to shutdown gracefully

You should see:
```
[INFO] Tomcat server stopped
```

## Technology Stack

- **Java**: 17
- **Apache Tomcat**: 10.1.13 (embedded)
- **Jakarta Servlet API**: 5.0.0
- **JXLS**: 2.10.0 (Template processing)
- **Apache POI**: 4.1.2 (Excel manipulation)
- **Maven Cargo Plugin**: 1.10.10 (Server management)

## Troubleshooting

### Issue: Port 8080 Already in Use

**Error**:
```
Address already in use: bind
```

**Solutions**:
1. Stop any other application using port 8080
2. Or modify the port in `pom.xml`:
   ```xml
   <cargo.servlet.port>9090</cargo.servlet.port>
   ```
   Then access at: `http://localhost:9090/excelgen/`

### Issue: Build Failures

**Error**:
```
There are test failures
```

**Solution**:
Always use `-DskipTests` flag:
```bash
mvn clean package -DskipTests
```

### Issue: Template Dropdown Shows "Error loading templates"

**Cause**: TemplateListServlet cannot find template files in resources

**Solution**:
1. Verify templates exist in `src/main/resources/`
2. Rebuild the project:
   ```bash
   mvn clean package -DskipTests
   ```
3. Restart the server

### Issue: "Missing required parameters" Error

**Cause**: Form data not being received correctly

**Solutions**:
1. Ensure you filled in all required fields (Name, Age, Template)
2. Check browser console for JavaScript errors
3. Verify server logs for detailed error messages

### Issue: Excel File Won't Open

**Cause**: Template processing error or invalid template

**Solution**:
1. Check server logs for JXLS errors
2. Verify the selected template is valid
3. Try a different template

### Issue: Java Version Error

**Error**:
```
java.lang.UnsupportedClassVersionError
```

**Solution**:
Ensure you're using Java 17 or higher:
```bash
java -version
```

If needed, set `JAVA_HOME`:
```bash
# Linux/Mac
export JAVA_HOME=/path/to/java17

# Windows
set JAVA_HOME=C:\path\to\java17
```

## Development Mode

### Enable Debug Logging

The servlet includes debug logging. Check console output for:
```
Received parameters:
  name: John Doe
  age: 25
  template: person_template.xlsx
  Content-Type: multipart/form-data
```

### Hot Reload

To apply code changes without restarting:
1. Stop the server (Ctrl+C)
2. Rebuild: `mvn package -DskipTests`
3. Start server: `mvn cargo:run`

## Production Deployment

### Build WAR File

```bash
mvn clean package
```

WAR file location: `target/excelgen-1.0-SNAPSHOT.war`

### Deploy to Standalone Tomcat

1. Copy WAR to Tomcat webapps directory:
   ```bash
   cp target/excelgen-1.0-SNAPSHOT.war /path/to/tomcat/webapps/
   ```

2. Start Tomcat:
   ```bash
   /path/to/tomcat/bin/startup.sh  # Linux/Mac
   /path/to/tomcat/bin/startup.bat # Windows
   ```

3. Access at: `http://localhost:8080/excelgen-1.0-SNAPSHOT/`

### Configure Context Path

To access at `/excelgen` instead of `/excelgen-1.0-SNAPSHOT`:

Rename the WAR file:
```bash
mv target/excelgen-1.0-SNAPSHOT.war target/excelgen.war
```

Or configure in Tomcat's `server.xml`.

## API Endpoints

If you want to use the servlets programmatically:

### List Templates
```bash
curl http://localhost:8080/excelgen/listTemplates
```

**Response**:
```json
["person_template.xlsx","person_template_address.xlsx","person_template_old_style.xlsx","person_template_with_phones.xlsx"]
```

### Process Template
```bash
curl -X POST http://localhost:8080/excelgen/processTemplate \
  -F name="John Doe" \
  -F age=25 \
  -F parentName="Jane Doe" \
  -F addressType="Home" \
  -F addressLine="123 Main St" \
  -F template="person_template.xlsx" \
  -o output.xls
```

**Response**: Excel file (binary)

### Process Template with Multiple Phones
```bash
curl -X POST http://localhost:8080/excelgen/processTemplate \
  -F name="Sarah Johnson" \
  -F age=28 \
  -F "phoneType[]=Mobile" \
  -F "phoneNo[]=+1-555-1234" \
  -F "phoneType[]=Office" \
  -F "phoneNo[]=+1-555-5678" \
  -F "phoneType[]=Home" \
  -F "phoneNo[]=+1-555-9012" \
  -F template="person_template_with_phones.xlsx" \
  -o output_phones.xls
```

**Response**: Excel file with multiple phone numbers (binary)

## Adding New Templates

To add your own templates:

1. Create JXLS 2.x template with proper commands in cell comments:
   ```
   Cell A1 comment: jx:area(lastCell="B10")
   Cell value: ${person.name}
   ```

2. Save as `.xlsx` file in `src/main/resources/`

3. Rebuild and restart:
   ```bash
   mvn clean package -DskipTests
   mvn cargo:run
   ```

4. Template will appear in the dropdown automatically

## Additional Resources

- [JXLS 2.x Documentation](http://jxls.sourceforge.net/)
- [Apache POI Documentation](https://poi.apache.org/)
- [Jakarta Servlet Specification](https://jakarta.ee/specifications/servlet/)
- [Main Project README](../CLAUDE.md)

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review server logs in the console
3. Verify all prerequisites are met
4. Check the main project documentation

## License

This project is for testing and educational purposes.
