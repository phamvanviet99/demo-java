package com.example.demo;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.Collections;

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) {
		SpringApplication app = new SpringApplication(DemoApplication.class);

		// Nếu Railway có PORT, lấy ra và set lại cho Tomcat
		String port = System.getenv("PORT");
		if (port != null) {
			app.setDefaultProperties(Collections.singletonMap("server.port", port));
			System.out.println("✅ Running on Railway port: " + port);
		} else {
			app.setDefaultProperties(Collections.singletonMap("server.port", "8080"));
			System.out.println("✅ Running on default port 8080");
		}

		app.run(args);
	}
}

