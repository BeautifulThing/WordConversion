package com.lin;


import com.lin.common.config.LibraryProperties;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.InitializingBean;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@Slf4j
public class WordChangeApplication implements InitializingBean {
	private final LibraryProperties library;

	public WordChangeApplication(LibraryProperties library) {
		this.library = library;
	}

	public static void main(String[] args) {
		SpringApplication.run(WordChangeApplication.class, args);

	}

	@Override
	public void afterPropertiesSet() throws Exception {
		log.info(library.getLocation());
		log.info(library.getBooks().toString());
		log.info(library.getDomain());
	}
}
