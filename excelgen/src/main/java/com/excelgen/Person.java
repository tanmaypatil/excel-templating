package com.excelgen;

import java.util.ArrayList;
import java.util.List;

public class Person {
    private String name;
    private int age;
    private String parentName;
    private Address address;
    private boolean addressExists;
    private List<Phone> phones;
    private boolean phoneExists;

    public Person() {
        this.addressExists = false;
        this.phones = new ArrayList<>();
        this.phoneExists = false;
    }

    public Person(String name, int age, String parentName) {
        this.name = name;
        this.age = age;
        this.parentName = parentName;
        this.addressExists = false;
        this.phones = new ArrayList<>();
        this.phoneExists = false;
    }

    public Person(String name, int age, String parentName, Address address) {
        this.name = name;
        this.age = age;
        this.parentName = parentName;
        this.address = address;
        this.addressExists = address != null;
        this.phones = new ArrayList<>();
        this.phoneExists = false;
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

    public List<Phone> getPhones() {
        return phones;
    }

    public void setPhones(List<Phone> phones) {
        this.phones = phones;
    }

    /**
     * Add a single phone number to the person's phone list
     */
    public void addPhone(Phone phone) {
        if (this.phones == null) {
            this.phones = new ArrayList<>();
        }
        this.phones.add(phone);
        this.phoneExists = true;
    }

    /**
     * Add a phone number with type and number
     */
    public void addPhone(String phoneType, String phoneNo) {
        if (this.phones == null) {
            this.phones = new ArrayList<>();
        }
        this.phones.add(new Phone(phoneType, phoneNo));
        this.phoneExists = true;
    }

    /**
     * Check if person has any phone numbers
     */
    public boolean hasPhones() {
        return phones != null && !phones.isEmpty();
    }

    public boolean isPhoneExists() {
        return phoneExists;
    }

    public void setPhoneExists(boolean phoneExists) {
        this.phoneExists = phoneExists;
    }
}
