



#You are on Development
- When new chat check machine & os to ensure proper understanding of environment 

### Development Machine (Windows) in WSL System
- Windows 11 (64-bit)
- Python 3.8+ (3.11 recommended)
- Node.js 18+ and npm 9+
- MySQL 8.0 Community Server
- Git for Windows
- 64GB RAM 
- 900GB free disk space






##File Size, Methods & Code Architecture (AI-Optimized)
- Refactor when limits are breached
- File Size Limits

--Core logic files: 150-200 lines
--Type definitions: 100 lines
--Single-purpose services: 200 lines
--Utilities: 150 lines per file
--Tests: 300 lines (more linear flow)

--Method Size Limit
---30 lines maximum - Extract method when exceeded

###Architecture Patterns

- Single Concept Per File - One clear purpose
- Predictable Structures - Consistent patterns across codebase
- Clear Boundaries - Explicit separation of concerns
- Code Organization & Refactoring
- Single Responsibility - Each function/class does one thing well
- No Code Duplication - Extract shared logic to utility modules/classes
- Utility Consolidation - Group related utilities.  Utilities are in their own directory.
- No harcoding, use variables; 
**Examples
/utils/dateHelpers.ts
/utils/stringHelpers.ts
/utils/sqliteHelpers.ts

- Shared Types - Centralize in /types or /models
- Constants - Single source of truth in /constants
- No Magic Numbers - Extract to named constants
- Extract Common Patterns - Database queries, validation logic, formatting functions

Approved Technology Stack & Tooling
C# backend / whatever is needed to convert the charts front end
UI: WPF



##PowerShell Integration

Scripts: Use PowerShell 7+ features
Error Handling: Use try-catch with proper error streams

##Performance Considerations

Caching: Use in-memory caching for frequently accessed data
Batch Operations: Use SQLite transactions for bulk operations
Streaming: Use async iterators for large datasets

# 7. Coding & Configuration Standards
* **Secrets**: store all keys & connection strings in `.env`; load via DotNetEnv.  
* **Logging**: use `ILogger<T>` injection; default level = `Information`; enrich logs with contextual properties.  
* **Error Handling**: catch at repository layer, wrap with custom exception, log with Serilog, bubble upward.  
* **Naming**: `PascalCase` for classes, `camelCase` for variables, `UPPER_SNAKE` for constants, `IXxx` for interfaces.  
* **Comments**: prefer self‑explanatory code; use XML doc comments for public APIs only.    
* **Internationalisation**: UI text through resource files; no hard‑coded strings.

## Security & Compliance
1. **Authentication & Authorization**: JWT bearer tokens, role‑based access (`Admin`, `User`).  
2. **OWASP**: Address Top‑10 items; mandatory threat‑model checklist per feature.  
3. **Data Privacy**: Mask or hash PII in logs; comply with GDPR/CCPA as applicable.  
4. **Encryption**: TLS 1.2+ for all network calls; AES‑256 at rest where feasible.  
5. **Audit Logging**: all admin actions and schema changes must be logged with user‑ID, timestamp, before/after values.

The application uses three distinct users with minimal required privileges:

### Disable dangerous features
local_infile = 0
secure_file_priv = NULL

### Enable audit logging
log_bin = mysql-bin
expire_logs_days = 7
binlog_format = ROW


### Performance with security
skip-name-resolve = 1




## Environment Configuration

### Configuration Files Structure
- Use best git best practice and web application best practice for folder structure.  We are to be clean and organized.
- Documentation should be in its own subfolder.  
- We do not propogate file after file. At most we can have 4-6 files that contain logical groupings of the following information **All approaches should follow industry best practices.  
- Documentation should detail: 
--App functionality 
--Module implementation flow
--Api flow 
--Security implementaiton (without providing a guide to hackers), 
--Virtual Environments and container implementation  
--Deployment guide including CI/CD Guide & Docker guide
### Documentation Standards
* Single `/.Docs` folder.  
* Update **existing** files; do not duplicate.  
* Minimum docs per feature: purpose, API surface, test coverage, security impact.  
* Diagrams as `.drawio` or `.png` committed to repo.


##Utilities
-Security Tools and Scripts
--Quick Security Status

DevOps & Git Workflow & Docker
1. **Branching**: `main` (prod), `develop` (integration), `feature/*`, `hotfix/*`.  
2. **Commits**: Conventional Commits style (`feat:`, `fix:`, `docs:`…).  
3. **Hooks**: pre‑commit checks run `dotnet format`, tests, and static analyzers.  
4. **CI/CD**: GitHub Actions pipeline stub; deploys via release tags.  
5. **No Auto‑Push**: scripts echo git commands; user confirms before push.
6. **Docker Supported**: Code should be implemented such that the application can run in docker our outside of docker 

## Performance & Monitoring
* Use `MemoryCache` for schema metadata (TTL 1 h).  
* Channel streaming for large data sets (avoids UI freeze).
	
## References
* Clean Code (Robert C. Martin)  
* OWASP Top 10 – 2025 Edition  
* Microsoft .NET Secure Coding Guidelines  