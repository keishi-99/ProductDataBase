package com.productdatabase.webviewer.auth;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.util.List;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.web.context.SecurityContextRepository;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/auth")
public class AuthController {

    private final AuthService authService;
    private final SecurityContextRepository securityContextRepository;

    public AuthController(AuthService authService, SecurityContextRepository securityContextRepository) {
        this.authService = authService;
        this.securityContextRepository = securityContextRepository;
    }

    @PostMapping("/login")
    public ResponseEntity<Void> login(
        @RequestBody LoginRequest request,
        HttpServletRequest httpRequest,
        HttpServletResponse httpResponse) {

        if (!authService.authenticate(request.password())) {
            return ResponseEntity.status(HttpStatus.UNAUTHORIZED).build();
        }

        var authentication = new UsernamePasswordAuthenticationToken("管理者", null, List.of());
        var context = SecurityContextHolder.createEmptyContext();
        context.setAuthentication(authentication);
        SecurityContextHolder.setContext(context);
        // Spring Security 6以降、SecurityContextHolderへのセットだけではセッションに保存されないため、
        // リポジトリへの保存を明示的に行う必要がある
        securityContextRepository.saveContext(context, httpRequest, httpResponse);

        return ResponseEntity.ok().build();
    }

    @PostMapping("/logout")
    public ResponseEntity<Void> logout(HttpServletRequest httpRequest) {
        httpRequest.getSession().invalidate();
        SecurityContextHolder.clearContext();
        return ResponseEntity.ok().build();
    }

    record LoginRequest(String password) {}
}
