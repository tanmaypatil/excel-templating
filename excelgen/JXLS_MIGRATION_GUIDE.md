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

## Syntax Comparison

### JXLS 1.x (Old - XML Style)
```xml
<jx:if test="person.age < 18">
    Parent: ${person.parentName}
</jx:if>
```

### JXLS 2.x (New - Command Style)
```
jx:if(condition="person.age < 18", areas=["A6:B6"])
```

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

#### Old Template (JXLS 1.x):
```
Row 1: Name: ${person.name}
Row 2: Age: ${person.age}
Row 3: <jx:if test="person.age < 18">
Row 4: Parent: ${person.parentName}
Row 5: </jx:if>
```

#### New Template (JXLS 2.x):
```
Row 1: Name: ${person.name}
Row 2: Age: ${person.age}
Row 3: jx:if(condition="person.age < 18", lastCell="B4", areas=["A4:B4"])
Row 4: Parent: ${person.parentName}
```

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

1. **Use Cell Comments for Commands** (keeps template clean)
2. **Always specify areas explicitly** (more predictable)
3. **Use descriptive condition expressions** (easier to debug)
4. **Test with edge cases** (null values, boundary conditions)

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

#### 1. Template Not Processing (Variables Appear as ${...})
**Problem**: Variables like `${person.name}` appear as literal text in output.

**Cause**: No valid JXLS 2.x commands in template, or invalid old JXLS 1.x syntax present.

**Solution**:
- Ensure at least one valid JXLS 2.x command exists
- Remove all old `<jx:if>` XML-style tags
- Use `jx:if(condition="...", lastCell="...")` instead

**Reference**: See `NegativeTemplateTest.java` in this project

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
