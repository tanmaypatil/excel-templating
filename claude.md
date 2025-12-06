# Excel Templating with Apache POI and JXLS 2.0

## Project Overview

This project demonstrates the use of Apache POI and JXLS 2.0 for generating Excel reports from templates with dynamic content and conditional logic. The project showcases how to create Excel templates programmatically and process them using JXLS expression language.

## Technology Stack

- **Java**: 17
- **Apache POI**: 4.1.2 (Excel file manipulation)
- **JXLS**: 2.10.0 (Template processing)
- **JXLS-POI**: 2.10.0 (POI integration for JXLS)
- **Apache Commons JEXL**: 2.1.1 (Expression evaluation)
- **JUnit**: 5.9.3 (Testing)
- **Maven**: 3.6.3+ (Build tool)

## Key Dependencies

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>
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
```

## Project Structure

```
excelgen/
├── src/
│   ├── main/
│   │   ├── java/com/excelgen/
│   │   │   ├── Person.java                      # Domain model for person data
│   │   │   ├── Address.java                     # Domain model for address data
│   │   │   ├── TemplateGenerator.java           # Template creation utility
│   │   │   ├── TemplateCreator.java             # Alternative template creator
│   │   │   ├── ImprovedTemplateGenerator.java   # Enhanced template generation
│   │   │   ├── OldStyleTemplateGenerator.java   # Legacy template generation
│   │   │   └── AddressTemplateGenerator.java    # Address template generator
│   │   └── resources/
│   │       ├── person_template.xlsx             # Static JXLS template (basic)
│   │       └── person_template_address.xlsx     # Static JXLS template (with address)
│   └── test/
│       └── java/com/excelgen/
│           ├── PersonTemplateTest.java          # Main template processing tests
│           ├── JxlsTemplateTest.java            # JXLS-specific tests
│           ├── NegativeTemplateTest.java        # Negative/edge case tests
│           └── VariableEvaluationTest.java      # Variable evaluation tests
└── pom.xml
```

## Core Concepts

### 1. Domain Models

#### Person
```java
public class Person {
    private String name;
    private int age;
    private String parentName;
    private Address address;
    private boolean addressExists;
}
```

#### Address
```java
public class Address {
    private String type;
    private String addressLine;
}
```

### 2. Template Creation

Templates can be created programmatically using Apache POI:

```java
Workbook workbook = new XSSFWorkbook();
Sheet sheet = workbook.createSheet("PersonInfo");

// Add title
Row row0 = sheet.createRow(0);
row0.createCell(0).setCellValue("Person Report");

// Add JXLS expressions
Row row2 = sheet.createRow(2);
row2.createCell(0).setCellValue("Name: ${person.name}");

Row row3 = sheet.createRow(3);
row3.createCell(0).setCellValue("Age: ${person.age}");

// Add conditional logic
Row row4 = sheet.createRow(4);
row4.createCell(0).setCellValue(
    "jx:if(condition=\"person.age < 18\" lastCell=\"A5\" areas=[\"A5:A5\"])"
);

Row row5 = sheet.createRow(5);
row5.createCell(0).setCellValue("Parent: ${person.parentName}");
```

### 3. Template Processing

Process templates using JxlsHelper:

```java
try (InputStream templateStream = getClass().getResourceAsStream("/person_template.xlsx");
     OutputStream outputStream = new FileOutputStream(outputPath)) {

    // Create JXLS context and bind data
    Context context = new Context();
    context.putVar("person", person);

    // Process template
    JxlsHelper.getInstance().processTemplate(templateStream, outputStream, context);
}
```

### 4. JXLS Conditional Logic

The templates use `jx:if` commands for conditional rendering:

```
jx:if(condition="person.age < 18" lastCell="A5" areas=["A5:A5"])
```

This command:
- Evaluates the condition `person.age < 18`
- If true, includes the specified areas (row with parent name)
- If false, excludes those areas from the output

## Use Cases

### Use Case 1: Minor (Age < 18)
When processing a person with age < 18, the output includes:
- Name
- Age
- Parent Name (conditional - shown)

### Use Case 2: Adult (Age >= 18)
When processing a person with age >= 18, the output includes:
- Name
- Age
- Parent Name (conditional - hidden)

### Use Case 3: Person with Address
When processing a person with an address, the output includes:
- Name
- Age
- Parent Name (conditional)
- Address Type
- Address Line

## Running the Project

### Build
```bash
cd excelgen
mvn clean install
```

### Run Tests
```bash
mvn test
```

### Generate Template
```bash
mvn exec:java -Dexec.mainClass="com.excelgen.TemplateGenerator"
```

## Test Cases

The project includes comprehensive test coverage:

1. **testProcessTemplateForMinor** - Validates template processing for minors
2. **testProcessTemplateForAdult** - Validates template processing for adults
3. **testProcessTemplateWithAddress** - Tests address integration
4. **testMinorIncludesParentName** - Verifies parent name appears for minors
5. **testAdultExcludesParentName** - Verifies parent name hidden for adults
6. **testBoundaryCase_Age18** - Tests edge case (exactly 18 years old)

## JXLS Expression Language

### Variable Access
```
${person.name}
${person.age}
${person.address.type}
```

### CRITICAL: Commands Must Be in Excel Cell COMMENTS

**IMPORTANT**: JXLS 2.x commands like `jx:area` and `jx:if` MUST be placed in Excel cell **comments**, not in cell values!

#### Why Cell Comments?
- JXLS 2.x uses `XlsCommentAreaBuilder` to read commands from cell comments
- Without valid commands in comments, expressions like `${person.name}` won't be replaced
- Cell values should contain display text and expressions, not commands

#### How to Add Commands in Cell Comments (Programmatically)

```java
// Create drawing patriarch for comments
Drawing<?> drawing = sheet.createDrawingPatriarch();
CreationHelper factory = workbook.getCreationHelper();

// Add jx:area command to cell A1 as a comment
Row row0 = sheet.createRow(0);
Cell cell = row0.createCell(0);
cell.setCellValue("Title");

ClientAnchor anchor = factory.createClientAnchor();
anchor.setCol1(0);
anchor.setCol2(3);
anchor.setRow1(0);
anchor.setRow2(2);
Comment comment = drawing.createCellComment(anchor);
comment.setString(factory.createRichTextString("jx:area(lastCell=\"B10\")"));
cell.setCellComment(comment);
```

#### How to Add Commands in Excel (Manually)
1. Right-click on the cell
2. Select "Insert Comment" or "New Note"
3. Type the JXLS command: `jx:area(lastCell="B10")`
4. Close the comment
5. Hover over the cell to see the red triangle indicator

### Conditional Rendering
```
jx:if(condition="person.age < 18", lastCell="B5")
```

### Required Command: jx:area
Every JXLS template must have a `jx:area` command in cell A1's comment:
```
jx:area(lastCell="B10")
```
This defines the processing area from A1 to B10.

### Parameters
- **condition**: Boolean expression to evaluate
- **lastCell**: Last cell in the command area
- **areas**: Cell ranges affected by the condition (alternative to lastCell)

## Key Features

1. **Dynamic Content Injection**: Replace placeholders with actual data
2. **Conditional Logic**: Show/hide sections based on conditions
3. **Nested Objects**: Access nested properties (person.address.type)
4. **Template Reusability**: Create templates once, use multiple times
5. **Type Safety**: Strongly-typed Java objects for data binding

## Common Patterns

### Pattern 1: Static Template
1. Create template file with JXLS expressions
2. Store in `src/main/resources/`
3. Load as resource stream
4. Process with JxlsHelper

### Pattern 2: Programmatic Template
1. Use Apache POI to create workbook
2. Add cells with JXLS expressions
3. Save as template file
4. Use like static template

## Troubleshooting

### Expressions Not Replaced (Shows ${...} in Output)
**Problem**: Variables like `${person.name}` appear as literal text in the output file.

**Cause**: JXLS commands are NOT in cell comments, or `jx:area` command is missing.

**Solution**:
1. **Add `jx:area` command** in cell A1's comment:
   ```
   jx:area(lastCell="B10")
   ```
2. **Put `jx:if` commands in cell COMMENTS**, not cell values:
   - Right-click cell → Insert Comment
   - Type: `jx:if(condition="person.age < 18", lastCell="B5")`
3. **Cell values** should contain only display text and `${...}` expressions
4. **Regenerate the template** if you modified it manually

**Reference**: See `AddressTemplateGenerator.java:40-48` for proper implementation

### Template Not Found
- Ensure template is in `src/main/resources/`
- Use correct resource path (e.g., `/person_template.xlsx`)

### Expressions Not Evaluated
- Verify JXLS context variable names match template expressions
- Check JEXL syntax in conditions
- **Most Common**: Commands must be in cell comments, not cell values!

### Conditional Not Working
- Validate `jx:if` command syntax in cell **comment**
- Ensure `lastCell` and `areas` parameters are correct
- Check condition expression returns boolean
- Verify the command is in a cell comment, not a cell value

## Best Practices

1. **Separate Template Creation from Processing**: Use dedicated generator classes
2. **Use Static Templates**: Prefer pre-created templates over runtime generation
3. **Test Boundary Cases**: Test edge conditions (age = 18, null values, etc.)
4. **Meaningful Variable Names**: Use descriptive context variable names
5. **Template Validation**: Open generated templates in Excel to verify structure

## Future Enhancements

Potential areas for expansion:
- Collection iteration (jx:each)
- Formula support
- Multiple sheet templates
- Custom JEXL functions
- Template inheritance
- Advanced formatting (colors, fonts, borders)
- Image embedding
- Chart generation

## Resources

- [Apache POI Documentation](https://poi.apache.org/)
- [JXLS Documentation](http://jxls.sourceforge.net/)
- [JEXL Expression Language](https://commons.apache.org/proper/commons-jexl/)

## License

This project is for testing and educational purposes.

## Notes

- Code generated with assistance from Claude Code
- Demonstrates integration of Apache POI and JXLS 2.0
- Focuses on conditional template rendering
- Uses JUnit 5 for testing
