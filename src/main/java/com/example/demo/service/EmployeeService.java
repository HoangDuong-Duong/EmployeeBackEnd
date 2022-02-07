package com.example.demo.service;

import com.example.demo.exception.UserNotFoundException;
import com.example.demo.model.Employee;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import com.example.demo.repo.EmployeeRepository;

import java.util.List;
import java.util.UUID;

@Service
public class EmployeeService {

    private final EmployeeRepository employeeRepository;

    @Autowired
    public EmployeeService(EmployeeRepository employeeRepository) {
        this.employeeRepository = employeeRepository;
    }
    public Employee addEmployee(Employee employee){
        employee.setEmployeeCode(UUID.randomUUID().toString());
        return employeeRepository.save(employee);
    }

    public List<Employee> findAllEmployee(){
        return employeeRepository.findAll();
    }

    public Employee updateEmployee(Employee employee){
        return  employeeRepository.save(employee);
    }

    public void deleteEmpoyee(Long id){
        employeeRepository.deleteById(id);
    }

    public Employee findEmployeeById(Long id){
        return employeeRepository.findEmployeeById(id).
                orElseThrow(()->new UserNotFoundException("User by id" + id+" was not found"));
    }
}
