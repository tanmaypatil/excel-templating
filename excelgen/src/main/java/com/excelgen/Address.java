package com.excelgen;

public class Address {
    private String type;
    private String addressLine;

    public Address() {
    }

    public Address(String type, String addressLine) {
        this.type = type;
        this.addressLine = addressLine;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getAddressLine() {
        return addressLine;
    }

    public void setAddressLine(String addressLine) {
        this.addressLine = addressLine;
    }
}
