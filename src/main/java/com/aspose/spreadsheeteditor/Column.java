/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.aspose.spreadsheeteditor;

import java.io.Serializable;
import javax.annotation.PostConstruct;

/**
 *
 * @author saqib
 */
public class Column implements Serializable {

    private int id;
    private String name;
    private String header;
    private String property;

    public Column(int id, String name) {
        this.id = id;
        this.name = this.header = this.property = name;
    }

    @PostConstruct
    public void init() {
    }

    public String getHeader() {
        return header;
    }

    public String getProperty() {
        return property;
    }

    public int getId() {
        return id;
    }

    public String getName() {
        return name;
    }

}
