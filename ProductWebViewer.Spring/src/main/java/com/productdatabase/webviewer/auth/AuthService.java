package com.productdatabase.webviewer.auth;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.util.List;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Service;

@Service
public class AuthService {

    @Value("${app.auth.admin-password}")
    private String adminPassword;

    public boolean authenticate(String submittedPassword) {
        if (submittedPassword == null) return false;

        var matches = MessageDigest.isEqual(
            submittedPassword.getBytes(StandardCharsets.UTF_8),
            adminPassword.getBytes(StandardCharsets.UTF_8)
        );
        if (!matches) return false;

        var authentication = new UsernamePasswordAuthenticationToken("管理者", null, List.of());
        SecurityContextHolder.getContext().setAuthentication(authentication);
        return true;
    }
}
