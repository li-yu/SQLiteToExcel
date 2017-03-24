package com.liyu.demo;

import org.litepal.crud.DataSupport;

/**
 * Created by liyu on 2017/3/24.
 */

public class User extends DataSupport {

    private String name;

    private float price;

    private byte[] cover;

    public byte[] getCover() {
        return cover;
    }

    public void setCover(byte[] cover) {
        this.cover = cover;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public float getPrice() {
        return price;
    }

    public void setPrice(float price) {
        this.price = price;
    }
}
