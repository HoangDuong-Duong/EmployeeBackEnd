package com.example.demo.Student;


import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

@Component
public class StudentDaoService {

    private static List<Student> studentList = new ArrayList<>();

    static{
        studentList.add(new Student(1,"Hoàng Dương",new Date()));
        studentList.add(new Student(2,"Jessica",new Date()));
        studentList.add(new Student(3,"Van der Beek",new Date()));
    }

    private static int UserCount = 3;

    public List<Student>retriveAll(){
        return studentList;
    }
    public Student save(Student student){
         if(student.getId() == 0){
             student.setId(++UserCount);
         }
         if(student.getBirthdate() == null){
             student.setBirthdate(new Date());
         }
         studentList.add(student);
         return student;
    }
    public Student findOne(int id){
         for(Student student : studentList){
             if(student.getId() == id){
                 return student;
             }
         }
         return null;
    }

    public Student deleteStudent(int id){
        Iterator<Student> studentIterator = studentList.iterator();
        while(studentIterator.hasNext()){
            Student student =  studentIterator.next();
            if(student.getId() == id){
                studentIterator.remove();
                return  student;
            }
        }
        return null;
    }


}
