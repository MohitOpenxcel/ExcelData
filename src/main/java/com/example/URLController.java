package com.example;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class URLController {

    @RequestMapping(value = "/" )
    public String handler (Model model) {
        model.addAttribute("msg",
                "a jar packaging example");
        return "index";
    }
}
