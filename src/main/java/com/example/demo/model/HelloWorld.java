package com.example.demo.model;



public class HelloWorld {
    private String message;

    public HelloWorld(String message){
        this.message = message;
    }

    public String getMessage(){
        return this.message;
    }
    public void setMessage(String message){
        this.message = message;
    }
}
