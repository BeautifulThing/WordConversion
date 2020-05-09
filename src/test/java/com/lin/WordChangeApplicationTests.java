package com.lin;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

@SpringBootTest
@Slf4j
class WordChangeApplicationTests {

	@Test
	void contextLoads() {
		String bunth = "#*#*0#*#*";
		String s = bunth.replaceAll("[#*]+'+1+'+[#*]+", "");
		System.out.println(s);
	}
}
