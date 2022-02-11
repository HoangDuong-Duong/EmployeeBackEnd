package com.example.demo.Student;

import java.util.Date;

public class ExceptionResponse {
    private Date timeStamp;
    private String massage;
    private String detail;

    public ExceptionResponse(Date timeStamp, String massage, String detail) {
        super();
        this.timeStamp = timeStamp;
        this.massage = massage;
        this.detail = detail;
    }

    public Date getTimeStamp() {
        return timeStamp;
    }

    public void setTimeStamp(Date timeStamp) {
        this.timeStamp = timeStamp;
    }

    public String getMassage() {
        return massage;
    }

    public void setMassage(String massage) {
        this.massage = massage;
    }

    public String getDetail() {
        return detail;
    }

    public void setDetail(String detail) {
        this.detail = detail;
    }
}
