# JXLS 1.x to 2.x Migration Guide

## Overview
JXLS 2.x completely changed the command syntax from XML-style tags to function-style commands.

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
