package com.productdatabase.webviewer.auth;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

@Service
public class AuthService {

    @Value("${app.auth.admin-password}")
    private String adminPassword;

    public boolean authenticate(String submittedPassword) {
        if (submittedPassword == null) return false;

        return MessageDigest.isEqual(
            submittedPassword.getBytes(StandardCharsets.UTF_8),
            adminPassword.getBytes(StandardCharsets.UTF_8)
        );
    }
}
