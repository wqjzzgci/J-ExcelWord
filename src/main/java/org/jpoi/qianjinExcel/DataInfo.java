package org.jpoi.qianjinExcel;

import java.util.List;
import java.util.Map;

/**
 * Created by wangqianjin on 2017/7/12.
 */
public class DataInfo {
    public String getClassCode() {
        return classCode;
    }

    public void setClassCode(String classCode) {
        this.classCode = classCode;
    }

    public String getFunctionCode() {
        return functionCode;
    }

    public void setFunctionCode(String functionCode) {
        this.functionCode = functionCode;
    }

    public String getUnitCode() {
        return unitCode;
    }

    public void setUnitCode(String unitCode) {
        this.unitCode = unitCode;
    }

    public String getUsercaseCode() {
        return usercaseCode;
    }

    public void setUsercaseCode(String usercaseCode) {
        this.usercaseCode = usercaseCode;
    }

    private String classCode;
    private String functionCode;
    private String unitCode;
    private String usercaseCode;


    public String getDaVersion() {
        return daVersion;
    }

    public void setDaVersion(String daVersion) {
        this.daVersion = daVersion;
    }

    private String daVersion;


    public Map<String, Integer> getUcollumn() {
        return ucollumn;
    }

    public void setUcollumn(Map<String, Integer> ucollumn) {
        this.ucollumn = ucollumn;
    }

    private Map<String,Integer> ucollumn;
}
