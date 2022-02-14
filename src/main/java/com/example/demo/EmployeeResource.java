package com.example.demo;

import com.example.demo.model.Employee;
import com.example.demo.model.HelloWorld;
import com.example.demo.service.EmployeeService;
import com.example.demo.util.ExcelExporter;
import com.example.demo.util.ExcelReportControl;
import com.example.demo.util.ExcelReportVNP;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

@RestController
@RequestMapping("/employee")
public class EmployeeResource {

    private final EmployeeService employeeService;

    @Autowired
    public EmployeeResource(EmployeeService employeeService) {
        this.employeeService = employeeService;
    }






    @GetMapping("/all")
    public ResponseEntity<List<Employee>> getAllEmployee(){
        List<Employee> employees = employeeService.findAllEmployee();
        return new ResponseEntity<>(employees, HttpStatus.OK);
    }

    @GetMapping("/find/{id}")
    public ResponseEntity<Employee> getEmployeeById(@PathVariable("id")Long id){
       Employee employee = employeeService.findEmployeeById(id);
       return new ResponseEntity<>(employee,HttpStatus.OK);
    }

    @PostMapping("/add")
    public ResponseEntity<Employee> addEmployee(@RequestBody Employee employee){
       Employee employee1 = employeeService.addEmployee(employee);
        return new ResponseEntity<>(employee1,HttpStatus.CREATED);
    }

    @PutMapping ("/update")
    public ResponseEntity<Employee> updateEmployee(@RequestBody Employee employee){
       Employee employee1 = employeeService.updateEmployee(employee);
        return new ResponseEntity<>(employee1,HttpStatus.OK);
    }

    @DeleteMapping("/delete/{id}")
    public ResponseEntity<?> deleteEmployee(@PathVariable("id")Long id){
        employeeService.deleteEmpoyee(id);
        return new ResponseEntity<>(HttpStatus.OK);
    }

    @GetMapping("/excel/general")
    public void exportToExcelGeneral(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue ="attachement; fileName =report.xlsx";
        response.setHeader(headerKey, headerValue);
        List<Employee> employeeList = employeeService.findAllEmployee();
        ExcelExporter excelExport = new ExcelExporter(employeeList);
        excelExport.export(response);
    }

    @GetMapping("/excel/reportcontrol")
    public void exportToExcelReportControl(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue ="attachement; fileName =ReportControl.xlsx";
        response.setHeader(headerKey, headerValue);
        List<Employee> employeeList = employeeService.findAllEmployee();
        ExcelReportControl excelReportControl = new ExcelReportControl(employeeList);
        excelReportControl.exportExcelReportControl(response);
    }

    @GetMapping("/excel/reportvnp")
    public void exportToExcelReporVNP(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue ="attachement; fileName =ReportControl.xlsx";
        response.setHeader(headerKey, headerValue);
        List<Employee> employeeList = employeeService.findAllEmployee();
        ExcelReportVNP excelReportVNP = new ExcelReportVNP(employeeList);
        excelReportVNP.exportExcelReportVNP(response);
    }





    @GetMapping("/helloWorld")
    public HelloWorld getMessage(){
        return new HelloWorld("This is Sparta");
    }


    @GetMapping("/helloWorld/{name}")
        public HelloWorld getMessageWithPathVariable(@PathVariable String name){
            return new HelloWorld(String.format("HelloWorld , %s",name ));
        }


}
