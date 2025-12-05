# Excel Generation with JXLS

A Java 17 Maven project demonstrating Excel template processing using Apache POI and JXLS 2.x with conditional logic.

## Project Overview

This project provides a working example of:
- **Apache POI** for Excel file manipulation
- **JXLS 2.x** for template-based Excel generation
- **Conditional rendering** using `jx:if` commands
- **Static templates** stored in resources
- **JUnit 5 tests** demonstrating the functionality

## What Was Built

### 1. Project Setup
- ✅ Upgraded from Java 8 to Java 17
- ✅ Updated Maven plugins for Java 17 compatibility
- ✅ Removed JaCoCo (code coverage)
- ✅ Removed Checkstyle (code quality)
- ✅ Minimal, clean Maven configuration

### 2. Dependencies Added
- **Apache POI 4.1.2** - Excel file manipulation
  - poi
  - poi-ooxml
  - poi-ooxml-schemas
  - poi-scratchpad
  - poi-excelant
- **JXLS 2.10.0** - Template processing
  - jxls
  - jxls-poi
- **Apache Commons JEXL 2.1.1** - Expression evaluation
- **SLF4J Simple 1.7.36** - Logging

### 3. Core Components

#### Person Model (`Person.java`)
Simple POJO with fields:
- `name` - Person's name
- `age` - Person's age
- `parentName` - Parent's name (conditionally displayed)

#### Static JXLS Template (`person_template.xlsx`)
Template with conditional logic:
```
Row 1: "Person Report" (title)
Row 3: Name: ${person.name}
Row 4: Age: ${person.age}
Row 6: jx:if(condition="person.age < 18", lastCell="B6")
Row 7: Parent: ${person.parentName}  ← Only shown if age < 18
```

#### JUnit Tests (`PersonTemplateTest.java`)
Comprehensive test suite:
- ✅ Minor (age < 18) - shows parent name
- ✅ Adult (age ≥ 18) - hides parent name
- ✅ Boundary case (age = 18) - hides parent name

## Project Structure

```
excelgen/
├── pom.xml                          # Maven configuration
├── README.md                        # This file
├── JXLS_MIGRATION_GUIDE.md         # JXLS 1.x → 2.x migration guide
├── src/
│   ├── main/
│   │   ├── java/com/excelgen/
│   │   │   ├── App.java                    # Main demo application
│   │   │   ├── Person.java                 # Person model
│   │   │   ├── TemplateCreator.java        # Template processing utility
│   │   │   ├── TemplateGenerator.java      # Static template generator
│   │   │   └── ImprovedTemplateGenerator.java
│   │   └── resources/
│   │       └── person_template.xlsx        # ★ Static JXLS template
│   └── test/
│       └── java/com/excelgen/
│           ├── PersonTemplateTest.java     # ★ Main test suite
│           ├── JxlsTemplateTest.java       # Additional tests
│           └── AppTest.java                # Basic app test
```

## Quick Start

### Prerequisites
- Java 17+
- Maven 3.6.3+

### Build Project
```bash
mvn clean compile
```

### Run Tests
```bash
# Run all tests
mvn test

# Run specific test
mvn test -Dtest=PersonTemplateTest
```

### Run Demo Application
```bash
mvn exec:java -Dexec.mainClass="com.excelgen.App"
```

## How to Use JXLS Templates

### 1. Create/Edit Template
Open `src/main/resources/person_template.xlsx` in Excel to see the template structure.

### 2. Use JxlsHelper to Process Template
```java
// Load template from resources
InputStream template = getClass().getResourceAsStream("/person_template.xlsx");
OutputStream output = new FileOutputStream("output.xlsx");

// Create context and bind data
Context context = new Context();
context.putVar("person", new Person("John Doe", 15, "Jane Doe"));

// Process template
JxlsHelper.getInstance().processTemplate(template, output, context);
```

### 3. JXLS 2.x Conditional Syntax
```
jx:if(condition="person.age < 18", lastCell="B6")
```

**Parameters:**
- `condition` - Boolean expression to evaluate
- `lastCell` - Last cell in the command area
- `areas` - (Optional) Explicit cell ranges to conditionally render

## JXLS Syntax (Version 2.x)

### Basic Expressions
```
${person.name}           # Variable substitution
${person.age + 1}        # Arithmetic
${person.name.length()}  # Method calls
```

### Conditional Rendering
```
jx:if(condition="expression", lastCell="B5")
jx:if(condition="expression", areas=["A5:B5"])
```

### Operators in Conditions
```
condition="age < 18"              # Less than
condition="age >= 18"             # Greater than or equal
condition="age >= 18 && active"   # Logical AND
condition="age < 18 || vip"       # Logical OR
condition="!empty(email)"         # Not empty
```

## Migration from JXLS 1.x

If you're migrating from JXLS 1.x (XML-style syntax), see **[JXLS_MIGRATION_GUIDE.md](JXLS_MIGRATION_GUIDE.md)** for:
- Complete syntax comparison
- Migration examples
- Common patterns
- Best practices

**Quick comparison:**
- **Old (1.x)**: `<jx:if test="age < 18">text</jx:if>`
- **New (2.x)**: `jx:if(condition="age < 18", areas=["A5:B5"])`

## Key Files

### Template File
- **Location**: `src/main/resources/person_template.xlsx`
- **Type**: Excel file with JXLS expressions
- **Purpose**: Static template for generating person reports
- **Can be edited**: Yes, open in Excel to modify

### Main Test
- **File**: `src/test/java/com/excelgen/PersonTemplateTest.java`
- **Purpose**: Demonstrates template processing with different scenarios
- **Run**: `mvn test -Dtest=PersonTemplateTest`

### Template Generator
- **File**: `src/main/java/com/excelgen/ImprovedTemplateGenerator.java`
- **Purpose**: Creates the static template (run once or to regenerate)
- **Run**: `mvn exec:java -Dexec.mainClass="com.excelgen.ImprovedTemplateGenerator"`

## Dependencies

### Core Dependencies
```xml
<!-- Apache POI - Excel manipulation -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>4.1.2</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>

<!-- JXLS - Template processing -->
<dependency>
    <groupId>org.jxls</groupId>
    <artifactId>jxls</artifactId>
    <version>2.10.0</version>
</dependency>
<dependency>
    <groupId>org.jxls</groupId>
    <artifactId>jxls-poi</artifactId>
    <version>2.10.0</version>
</dependency>

<!-- Expression evaluation -->
<dependency>
    <groupId>org.apache.commons</groupId>
    <artifactId>commons-jexl</artifactId>
    <version>2.1.1</version>
</dependency>

<!-- Logging -->
<dependency>
    <groupId>org.slf4j</groupId>
    <artifactId>slf4j-simple</artifactId>
    <version>1.7.36</version>
</dependency>
```

### Test Dependencies
```xml
<dependency>
    <groupId>org.junit.jupiter</groupId>
    <artifactId>junit-jupiter-api</artifactId>
    <version>5.9.3</version>
    <scope>test</scope>
</dependency>
<dependency>
    <groupId>org.junit.jupiter</groupId>
    <artifactId>junit-jupiter-engine</artifactId>
    <version>5.9.3</version>
    <scope>test</scope>
</dependency>
```

## Configuration

### Java Version
- **Source**: Java 17
- **Target**: Java 17

### Maven Plugins
- **maven-compiler-plugin**: 3.13.0
- **maven-surefire-plugin**: 3.1.2 (for JUnit 5)
- **maven-enforcer-plugin**: 3.3.0 (enforces Maven 3.6.3+)
- **maven-javadoc-plugin**: 3.5.0

## Common Tasks

### Generate New Template
```bash
mvn exec:java -Dexec.mainClass="com.excelgen.ImprovedTemplateGenerator"
```

### Run All Tests
```bash
mvn test
```

### Run Specific Test Method
```bash
mvn test -Dtest=PersonTemplateTest#testMinorIncludesParentName
```

### Package Project
```bash
mvn package
```

### Clean Build
```bash
mvn clean install
```

## Testing

### Test Coverage
- ✅ Minor person (age < 18) shows parent name
- ✅ Adult person (age ≥ 18) hides parent name
- ✅ Boundary case (age = 18) hides parent name
- ✅ Template loading from resources
- ✅ JXLS expression evaluation
- ✅ Conditional rendering with jx:if

### Test Output
Tests generate Excel files in temporary directories:
```
/var/folders/.../junit.../minor_output.xlsx
/var/folders/.../junit.../adult_output.xlsx
/var/folders/.../junit.../boundary_18.xlsx
```

## Troubleshooting

### Template Not Found
**Error**: `Template not found: /person_template.xlsx`

**Solution**: Ensure template is in `src/main/resources/` and rebuild:
```bash
mvn clean compile
```

### JXLS Command Not Working
**Issue**: Conditional content always/never appears

**Check**:
1. Command syntax: `jx:if(condition="...", lastCell="...")`
2. Cell references are correct
3. Expression evaluates to boolean
4. Template is in resources and properly loaded

### Java Version Issues
**Error**: Class file version mismatch

**Solution**: Verify Java 17:
```bash
java -version  # Should show 17.x
mvn -version   # Should show Java 17
```

## Examples

### Example 1: Minor Person (Shows Parent)
```java
Person minor = new Person("John Doe", 15, "Jane Doe");
processTemplate("/person_template.xlsx", "output.xlsx", minor);
```

**Output Excel:**
```
Person Report

Name: John Doe
Age: 15

Parent: Jane Doe  ← Appears because age < 18
```

### Example 2: Adult Person (Hides Parent)
```java
Person adult = new Person("Alice Smith", 25, "Bob Smith");
processTemplate("/person_template.xlsx", "output.xlsx", adult);
```

**Output Excel:**
```
Person Report

Name: Alice Smith
Age: 25
                  ← Parent line not shown because age >= 18
```

## Resources

### Documentation
- [JXLS 2.x Documentation](http://jxls.sourceforge.net/)
- [Apache POI Documentation](https://poi.apache.org/)
- [JXLS Migration Guide](JXLS_MIGRATION_GUIDE.md) (in this project)

### Related Links
- [JXLS GitHub](https://github.com/jxlsteam/jxls)
- [Apache POI GitHub](https://github.com/apache/poi)

## Development History

### Changes Made
1. ✅ Upgraded from Java 8 to Java 17
2. ✅ Updated Maven plugins for Java 17 compatibility
3. ✅ Removed JaCoCo code coverage plugin
4. ✅ Removed Checkstyle plugin
5. ✅ Added Apache POI 4.1.2 dependencies
6. ✅ Added JXLS 2.10.0 dependencies
7. ✅ Added Commons JEXL 2.1.1
8. ✅ Added SLF4J Simple 1.7.36
9. ✅ Created Person model
10. ✅ Created static JXLS template with conditional logic
11. ✅ Created comprehensive JUnit test suite
12. ✅ Created JXLS migration guide

## License

This is a demonstration project for learning JXLS and Apache POI.

## Support

For issues or questions about:
- **JXLS syntax**: See [JXLS_MIGRATION_GUIDE.md](JXLS_MIGRATION_GUIDE.md)
- **Template issues**: Check `src/main/resources/person_template.xlsx`
- **Test failures**: Run `mvn test -X` for detailed output

---

**Version**: 1.0-SNAPSHOT
**Java**: 17
**JXLS**: 2.10.0
**Apache POI**: 4.1.2
