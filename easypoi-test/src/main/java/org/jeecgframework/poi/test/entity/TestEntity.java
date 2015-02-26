package org.jeecgframework.poi.test.entity;

import org.jeecgframework.poi.excel.annotation.Excel;

public class TestEntity {

    @Excel(name = "蓝牙POS机")
    private String lanya;
    @Excel(name = "蓝牙点付宝")
    private String pos;

    public String getLanya() {
        return lanya;
    }

    public void setLanya(String lanya) {
        this.lanya = lanya;
    }

    public String getPos() {
        return pos;
    }

    public void setPos(String pos) {
        this.pos = pos;
    }

}
