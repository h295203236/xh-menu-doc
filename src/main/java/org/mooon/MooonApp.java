package org.mooon;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import java.io.IOException;
import java.text.ParseException;

@SpringBootApplication
public class MooonApp {
    public static void main(String[] args) throws IOException, ParseException {
        SpringApplication.run(MooonApp.class, args);
    }
}
