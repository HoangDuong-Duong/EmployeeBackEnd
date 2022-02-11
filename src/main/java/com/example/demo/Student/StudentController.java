package com.example.demo.Student;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.hateoas.EntityModel;
import org.springframework.hateoas.server.mvc.WebMvcLinkBuilder;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import javax.validation.Valid;
import java.util.List;

import static org.springframework.hateoas.server.mvc.WebMvcLinkBuilder.linkTo;
import static org.springframework.hateoas.server.mvc.WebMvcLinkBuilder.methodOn;

@RestController
public class StudentController {
    @Autowired
    private StudentDaoService daoService;


    @GetMapping("/all")
    public List<Student> restriveAll() {
        return daoService.retriveAll();
    }

    @GetMapping("/find/{id}")
    public EntityModel<Student> findOne(@PathVariable int id) {
        Student student =  daoService.findOne(id);
        if(student == null){
            throw new StudentNotFoundException("id = "+id );
        }
        EntityModel<Student> model =   EntityModel.of(student);
        WebMvcLinkBuilder linktoAllUser = linkTo(methodOn(this.getClass()).restriveAll());
        model.add(linktoAllUser.withRel("all-user"));
        return model;
    }

    @PostMapping("/add")
    public ResponseEntity<Student> createUser(@Valid @RequestBody Student student) {
        Student student1 = daoService.save(student);
        return new ResponseEntity<>(student1, HttpStatus.CREATED);
    }

    @DeleteMapping("/delete/{id}")
    public ResponseEntity<Object> deleteStudent(@PathVariable int id){
        Student student = daoService.deleteStudent(id);
        if(student == null){
            throw new StudentNotFoundException("id= "+id);
        }
        return new ResponseEntity<>("Delete Success", HttpStatus.OK);
    }


}
