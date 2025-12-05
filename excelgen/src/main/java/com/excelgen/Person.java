package com.excelgen;

public class Person {
    private String name;
    private int age;
    private String parentName;

    public Person() {
    }

    public Person(String name, int age, String parentName) {
        this.name = name;
        this.age = age;
        this.parentName = parentName;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getParentName() {
        return parentName;
    }

    public void setParentName(String parentName) {
        this.parentName = parentName;
    }
}
