package com.excelgen;

public class Person {
    private String name;
    private int age;
    private String parentName;
    private Address address;
    private boolean addressExists;

    public Person() {
        this.addressExists = false;
    }

    public Person(String name, int age, String parentName) {
        this.name = name;
        this.age = age;
        this.parentName = parentName;
        this.addressExists = false;
    }

    public Person(String name, int age, String parentName, Address address) {
        this.name = name;
        this.age = age;
        this.parentName = parentName;
        this.address = address;
        this.addressExists = address != null;
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

    public Address getAddress() {
        return address;
    }

    public void setAddress(Address address) {
        this.address = address;
        this.addressExists = address != null;
    }

    public boolean isAddressExists() {
        return addressExists;
    }

    public void setAddressExists(boolean addressExists) {
        this.addressExists = addressExists;
    }
}
