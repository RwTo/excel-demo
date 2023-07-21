package com.rwto.excel.easyexcel;

import lombok.Data;

@Data
public class UserData {
    private String name;
    private int age;
    private String email;

    public UserData() {
    }

    public UserData(String name, int age, String email) {
        this.name = name;
        this.age = age;
        this.email = email;
    }
}