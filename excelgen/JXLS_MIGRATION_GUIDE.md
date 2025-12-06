# JXLS 1.x to 2.x Migration Guide

## Overview
JXLS 2.x completely changed the command syntax from XML-style tags to function-style commands.

## Official JXLS Documentation

### JXLS 2.x Official References:
- **Main Documentation**: https://jxls.sourceforge.net/
- **Getting Started Guide**: https://jxls.sourceforge.net/getting-started.html
- **If Command Reference**: https://jxls.sourceforge.net/reference/if_command.html
- **Each Command Reference**: https://jxls.sourceforge.net/reference/each_command.html
- **All Commands Overview**: https://jxls.sourceforge.net/reference/overview.html
- **Expression Language**: https://jxls.sourceforge.net/reference/expression_language.html
- **GitHub Repository**: https://github.com/jxlsteam/jxls

### Important Notes:
- This guide covers migration from **JXLS 1.x** (net.sf.jxls) to **JXLS 2.x** (org.jxls)
- Current project uses: **JXLS 2.10.0** with **Apache POI 4.1.2**
- Latest JXLS version: **3.0.0** (requires separate migration if upgrading)
- JXLS 2.x requires **Java 8+** (this project uses Java 17)

---

## ⚠️ CRITICAL REQUIREMENT FOR JXLS 2.x

### Commands MUST Be in Excel Cell Comments!

The **#1 most common mistake** when migrating to JXLS 2.x:

❌ **WRONG**: Putting commands as cell values
```
Cell A1 value: "jx:area(lastCell="B10")"
Cell A5 value: "jx:if(condition="person.age < 18", lastCell="B5")"
```
Result: Template won't process, expressions show as `${person.name}` literally

✅ **CORRECT**: Putting commands in cell comments
```
Cell A1 value: "Person Report"
Cell A1 comment: jx:area(lastCell="B10")

Cell A5 value: "Parent:"
Cell A5 comment: jx:if(condition="person.age < 18", lastCell="B5")
```
Result: Template processes correctly, expressions replaced with actual data

**How to add comments:**
- **In Excel**: Right-click cell → Insert Comment → Type command
- **In Java/POI**: Use `Drawing.createCellComment()` (see examples below)

**Why?** JXLS 2.x uses `XlsCommentAreaBuilder` which only reads commands from Excel cell comments, not from cell values.

---

## Syntax Comparison

### JXLS 1.x (Old - XML Style)
```xml
<jx:if test="person.age < 18">
    Parent: ${person.parentName}
</jx:if>
```

### JXLS 2.x (New - Command Style)

**CRITICAL**: Commands must be in Excel cell **COMMENTS**, not cell values!

```
jx:if(condition="person.age < 18", areas=["A6:B6"])
```

This command should be placed in an Excel cell comment (right-click → Insert Comment), NOT as a cell value.

## Migration Steps

### Step 1: Update Dependencies

**Remove JXLS 1.x:**
```xml
<dependency>
    <groupId>net.sf.jxls</groupId>
    <artifactId>jxls-core</artifactId>
    <version>1.0.6</version>
</dependency>
```

**Add JXLS 2.x:**
```xml
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

### Step 2: Update Template Syntax

#### Important: Commands Go in Cell Comments!

In JXLS 2.x, commands like `jx:area` and `jx:if` **MUST** be placed in Excel cell comments, not in cell values.

**How to add commands in Excel:**
1. Right-click on the cell
2. Select "Insert Comment" or "New Note"
3. Type the JXLS command (e.g., `jx:area(lastCell="B10")`)
4. Close the comment

**How to add commands programmatically (Java/POI):**
```java
Drawing<?> drawing = sheet.createDrawingPatriarch();
CreationHelper factory = workbook.getCreationHelper();

Cell cell = row.createCell(0);
cell.setCellValue("Parent:");

ClientAnchor anchor = factory.createClientAnchor();
anchor.setCol1(0);
anchor.setCol2(2);
anchor.setRow1(5);
anchor.setRow2(6);
Comment comment = drawing.createCellComment(anchor);
comment.setString(factory.createRichTextString("jx:if(condition=\"person.age < 18\", lastCell=\"B5\")"));
cell.setCellComment(comment);
```

#### Old Template (JXLS 1.x):
```
Row 1: Name: ${person.name}
Row 2: Age: ${person.age}
Row 3: <jx:if test="person.age < 18">
Row 4: Parent: ${person.parentName}
Row 5: </jx:if>
```

#### New Template (JXLS 2.x):

**Excel Structure:**
```
Row 1 (A1): "Person Report"
  └─ Cell Comment: jx:area(lastCell="B4")

Row 2: Name: | ${person.name}
Row 3: Age:  | ${person.age}

Row 4: Parent: | ${person.parentName}
  └─ Cell Comment on A4: jx:if(condition="person.age < 18", lastCell="B4")
```

**What you see in Excel cells:**
- Cell A1: "Person Report" (with comment indicator)
- Cell A2: "Name:", Cell B2: "${person.name}"
- Cell A3: "Age:", Cell B3: "${person.age}"
- Cell A4: "Parent:", Cell B4: "${person.parentName}" (with comment indicator)

**What's in the cell comments:**
- A1 comment: `jx:area(lastCell="B4")`
- A4 comment: `jx:if(condition="person.age < 18", lastCell="B4")`

### Step 3: Update Java Code

#### Old Code (JXLS 1.x):
```java
XLSTransformer transformer = new XLSTransformer();
Map<String, Object> beans = new HashMap<>();
beans.put("person", person);
transformer.transformXLS(templatePath, beans, outputPath);
```

#### New Code (JXLS 2.x):
```java
try (InputStream is = new FileInputStream(templatePath);
     OutputStream os = new FileOutputStream(outputPath)) {

    Context context = new Context();
    context.putVar("person", person);

    JxlsHelper.getInstance().processTemplate(is, os, context);
}
```

## Command Parameters Explained

### condition (required)
The boolean expression to evaluate:
```
condition="person.age < 18"
condition="person.age >= 18 && person.active"
condition="!empty(person.email)"
```

### areas (recommended)
Specifies which cell range to conditionally render:
```
areas=["A5:B5"]          // Single row
areas=["A5:B10"]         // Multiple rows
areas=["A5:B5", "A7:B7"] // Multiple separate areas
```

### lastCell (alternative to areas)
Specifies the last cell of the command area:
```
lastCell="B5"  // Command area ends at B5
```

## Common Migration Patterns

### Pattern 1: Simple Condition
**Old:**
```xml
<jx:if test="showDetails">
    Details: ${details}
</jx:if>
```

**New:**
```
jx:if(condition="showDetails", areas=["A2:B2"])
Details: ${details}
```

### Pattern 2: Nested Conditions
**Old:**
```xml
<jx:if test="person.age < 18">
    <jx:if test="person.hasGuardian">
        Guardian: ${person.guardian}
    </jx:if>
</jx:if>
```

**New:**
```
Row 1: jx:if(condition="person.age < 18", areas=["A2:B3"])
Row 2:   jx:if(condition="person.hasGuardian", areas=["A3:B3"])
Row 3:     Guardian: ${person.guardian}
```

### Pattern 3: Else Clause
**Old:**
```xml
<jx:if test="person.age < 18">
    Status: Minor
    <jx:else>
    Status: Adult
    </jx:else>
</jx:if>
```

**New:**
```
Row 1: jx:if(condition="person.age < 18", areas=["A2:B2"])
Row 2:   Status: Minor
Row 3: jx:if(condition="person.age >= 18", areas=["A4:B4"])
Row 4:   Status: Adult
```

Or using built-in else:
```
Row 1: jx:if(condition="person.age < 18", areas=["A2:B2"], ifArea="A2:B2", elseArea="A3:B3")
Row 2:   Status: Minor
Row 3:   Status: Adult
```

## Best Practices for JXLS 2.x

1. **ALWAYS Use Cell Comments for Commands** ⚠️ CRITICAL
   - Commands (`jx:area`, `jx:if`, `jx:each`) go in Excel cell comments
   - Cell values contain only display text and `${...}` expressions
   - This is NOT optional - JXLS 2.x requires this!
2. **Always add jx:area command** in cell A1's comment (required for processing)
3. **Always specify areas/lastCell explicitly** (more predictable)
4. **Use descriptive condition expressions** (easier to debug)
5. **Test with edge cases** (null values, boundary conditions)
6. **Open template in Excel** to verify comments are present (look for red triangles)

## Complete Working Example

See the `PersonTemplateTest.java` for a complete working example with:
- Static template in resources
- Conditional parent name display
- JUnit tests demonstrating the functionality

## Additional JXLS 2.x Commands

### Each Command (Loops)
**Documentation**: https://jxls.sourceforge.net/reference/each_command.html

```
jx:each(items="employees", var="employee", lastCell="E4")
```

### Grid Command (Tables)
**Documentation**: https://jxls.sourceforge.net/reference/grid_command.html

```
jx:grid(headers="headers", data="data", areas=["A1:C1","A2:C2"], formatCells="A2:C2")
```

### Image Command
**Documentation**: https://jxls.sourceforge.net/reference/image_command.html

```
jx:image(lastCell="B2", src="image")
```

## Troubleshooting

### Common Issues and Solutions

#### 1. Template Not Processing (Variables Appear as ${...}) - MOST COMMON!
**Problem**: Variables like `${person.name}` appear as literal text in output.

**Cause**: Commands are in cell VALUES instead of cell COMMENTS, or missing `jx:area` command.

**Solution**:
1. **Put commands in cell COMMENTS, not cell values!**
   - Right-click cell → Insert Comment
   - Type command in the comment (e.g., `jx:area(lastCell="B10")`)
2. **Add `jx:area` command** in cell A1's comment (required!)
3. **Cell values** should only contain:
   - Display text (e.g., "Name:", "Age:")
   - Expressions (e.g., `${person.name}`, `${person.age}`)
4. Remove all old `<jx:if>` XML-style tags
5. Regenerate template if you modified it manually

**Example - WRONG (command as cell value):**
```
Cell A5 value: "jx:if(condition="person.age < 18", lastCell="B5")"  ❌
```

**Example - CORRECT (command in cell comment):**
```
Cell A5 value: "Parent:"  ✅
Cell A5 comment: jx:if(condition="person.age < 18", lastCell="B5")  ✅
```

**Reference**: See `AddressTemplateGenerator.java` lines 40-48 for correct implementation

#### 2. Condition Not Working
**Problem**: Content always/never appears regardless of condition.

**Cause**:
- Incorrect condition syntax
- Wrong cell references in areas/lastCell
- Expression evaluation errors

**Solution**:
- Verify condition syntax: `condition="variable < 18"` (not `test="..."`)
- Check cell references match your content
- Test condition in isolation
- Enable debug logging (SLF4J)

**Reference**: https://jxls.sourceforge.net/reference/if_command.html

#### 3. ClassNotFoundException or NoSuchMethodError
**Problem**: Runtime errors when running JXLS code.

**Cause**: Mixing JXLS 1.x and 2.x dependencies.

**Solution**:
- Remove all `net.sf.jxls` dependencies
- Use only `org.jxls` dependencies
- Check transitive dependencies with `mvn dependency:tree`

#### 4. Old Syntax Still in Template
**Problem**: Forgot to update some templates.

**Solution**:
- Search project for `<jx:if` to find old syntax
- Use negative tests to verify (see `NegativeTemplateTest.java`)
- Run all tests after migration

## Learning Resources

### Official Documentation
1. **Quick Start**: https://jxls.sourceforge.net/getting-started.html
2. **Command Reference**: https://jxls.sourceforge.net/reference/overview.html
3. **Expression Language**: https://jxls.sourceforge.net/reference/expression_language.html

### Example Projects
- This project's tests: `src/test/java/com/excelgen/PersonTemplateTest.java`
- JXLS GitHub Examples: https://github.com/jxlsteam/jxls/tree/master/jxls-examples

### Community Support
- **GitHub Discussions**: https://github.com/jxlsteam/jxls/discussions
- **GitHub Issues**: https://github.com/jxlsteam/jxls/issues

## Version Information

### This Project:
- **JXLS**: 2.10.0
- **Apache POI**: 4.1.2
- **Java**: 17
- **JEXL**: 2.1.1

### Latest Versions:
- **JXLS**: 3.0.0 (requires migration from 2.x)
- **Apache POI**: 5.2.x (compatible with Java 8+)

### Upgrade Path:
1. ✅ JXLS 1.x → 2.x (this guide)
2. ⏭️ JXLS 2.x → 3.x (see: https://jxls.sourceforge.net/reference/migration_to_v3.html)

## Quick Reference Card

### JXLS 1.x vs 2.x Command Syntax

| Feature | JXLS 1.x | JXLS 2.x |
|---------|----------|----------|
| **If condition** | `<jx:if test="...">` | `jx:if(condition="...")` |
| **If closing** | `</jx:if>` | Not needed (use areas) |
| **Each loop** | `<jx:forEach items="..." var="...">` | `jx:each(items="..." var="...")` |
| **Each closing** | `</jx:forEach>` | Not needed (use areas) |
| **Variables** | `${variable}` | `${variable}` (same) |
| **Context** | `Map<String, Object>` | `Context` object |
| **API** | `XLSTransformer` | `JxlsHelper` |
| **Package** | `net.sf.jxls` | `org.jxls` |

### Parameter Names

| JXLS 1.x | JXLS 2.x | Notes |
|----------|----------|-------|
| `test=` | `condition=` | Boolean expression |
| `items=` | `items=` | Same |
| `var=` | `var=` | Same |
| N/A | `lastCell=` | Defines area end |
| N/A | `areas=` | Explicit area definition |

## See Also

- **README.md**: Project overview and setup
- **PersonTemplateTest.java**: Working examples
- **NegativeTemplateTest.java**: What NOT to do (old syntax examples)
- **VariableEvaluationTest.java**: Understanding when variables are evaluated

---

**Last Updated**: December 2025
**JXLS Version Covered**: 2.x (specifically 2.10.0)
**Official Documentation**: https://jxls.sourceforge.net/
