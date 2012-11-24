package test.colin.wielga;

import static org.junit.Assert.*;

import org.junit.Test;

import com.colin.wielga.LineEncoder;

public class LineEncodeCheck {
String[] shouldbe =	{"==","==","=="
    	,"==","==","(S","(I","PR","PR","PR","PR","==","I)","S)","(S","(I","PR","PR","PR"
    	,"==","I)","S)","(S","(I","PR","PR","PR","==","I)","S)","(S","S)","(S","(I","MS"
    	,"EE","MS","EE","MS","EE","MS","EE","MS","EE","MS","EE","MS","EE","MS","EE","MS"
    	,"EE","MS","EE","MS","I)","S)","(S","==","==","S)","(S","ED","S)","(S","==","S)"
    	,"(S","(I","PR","PR","PR","==","I)","S)","(S","(I","PR","PR","PR","==","I)","S)"
    	,"(S","(I","PR","PR","PR","==","I)","S)","(S","(I","PR","PR","PR","==","I)","S)"
    	,"(S","(I","PR","PR","PR","==","I)","S)","(S","(I","PR","PR","PR","==","I)","S)"
    	,"(S","(I","PR","PR","PR","==","I)","S)","(S","(I","PR","PR","PR","==","I)","S)"};
	@Test
	public void testEncode() {
		LineEncoder.load("lineencoding1");
	    assertEquals("Result",shouldbe, LineEncoder.encode("string"));
	}
}
