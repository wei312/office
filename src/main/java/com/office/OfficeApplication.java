/**
 * Copyright (c) 2016-2019 人人开源 All rights reserved.
 *
 * https://www.renren.io
 *
 * 版权所有，侵权必究！
 */

package com.office;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;


@SpringBootApplication
public class OfficeApplication extends SpringBootServletInitializer {

	@Override
	protected SpringApplicationBuilder configure(SpringApplicationBuilder application) {
		return application.sources(OfficeApplication.class);
	}

	public static void main(String[] args) {
		SpringApplication.run(OfficeApplication.class, args);
	}

}