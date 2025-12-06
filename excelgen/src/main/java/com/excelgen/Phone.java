package com.excelgen;

public class Phone {
    private String phoneType;
    private String phoneNo;

    public Phone() {
    }

    public Phone(String phoneType, String phoneNo) {
        this.phoneType = phoneType;
        this.phoneNo = phoneNo;
    }

    public String getPhoneType() {
        return phoneType;
    }

    public void setPhoneType(String phoneType) {
        this.phoneType = phoneType;
    }

    public String getPhoneNo() {
        return phoneNo;
    }

    public void setPhoneNo(String phoneNo) {
        this.phoneNo = phoneNo;
    }

    @Override
    public String toString() {
        return "Phone{" +
                "phoneType='" + phoneType + '\'' +
                ", phoneNo='" + phoneNo + '\'' +
                '}';
    }
}
