# Testing Scope and Strategy

## Overview

This document defines the testing scope and strategies for the Windows automation toolkit, particularly addressing the challenges of testing Office application automation.

## Testing Categories

### 1. Unit Testing (Fully Automated)

**Scope:** All utility functions and wrapper classes that abstract Windows API calls

**Coverage:**
- Windows API wrapper utilities (pywin32/pywinauto)
- Script discovery and loading mechanisms
- Configuration parsing and validation
- CLI argument processing
- Error handling utilities

**Approach:**
- Mock all Windows API calls using `pytest-mock`
- Mock file system operations when needed
- Focus on business logic rather than actual Windows interactions

**Mockable Components:**
```python
# Examples of what gets mocked
- win32api.GetSystemDirectory()
- pywinauto.Application()
- win32gui.FindWindow()
- win32clipboard operations
```

### 2. Integration Testing (Partially Automated)

**Scope:** User scripts that don't require specific applications

**Coverage:**
- Scripts that interact with basic Windows features (clipboard, notifications)
- Scripts that work with generic Windows controls
- File system operations
- Registry operations (read-only, safe operations)

**Requirements:**
- Windows environment (GitHub Actions Windows runner)
- No external applications required
- Safe operations only (no system modifications)

**Test Environment:**
- GitHub Actions: `windows-latest` runner
- Local development: Windows 10/11

### 3. Manual/Semi-Automated Testing (Office Applications)

**Scope:** Scripts that interact with Office applications

**Limitations:**
Office applications cannot be reliably automated in CI environments because:
- **Licensing**: Office requires valid licenses in CI
- **Interactive Elements**: Office apps may show dialogs, activation prompts
- **Environment Complexity**: COM registration, Office versions, updates
- **Resource Usage**: Office apps are resource-intensive for CI

**Recommended Approach:**

#### Option A: Mock-Based Testing
```python
# Mock Office COM objects
@patch('win32com.client.Dispatch')
def test_excel_automation(mock_dispatch):
    mock_excel = MagicMock()
    mock_dispatch.return_value = mock_excel
    # Test business logic without actual Excel
```

#### Option B: Local Testing Framework
- Provide testing templates for Office scripts
- Manual testing checklist
- Local test runner for Office-dependent scripts
- Test data fixtures and expected results

#### Option C: Hybrid Approach
- Unit test business logic with mocks
- Provide integration test templates for manual execution
- Document test cases and expected behaviors

### 4. Smoke Testing (Manual)

**Scope:** End-to-end validation of complete workflows

**Coverage:**
- Full script execution with real applications
- User workflow validation
- Performance testing
- Error handling in real scenarios

## Testing Strategy by Script Type

### Office Automation Scripts

**Unit Testing:**
- ✅ Mock all Office COM objects
- ✅ Test parameter validation
- ✅ Test error handling logic
- ✅ Test data processing functions

**Integration Testing:**
- ❌ Cannot reliably test in CI
- ✅ Local testing with test documents
- ✅ Template-based testing framework

**Recommended Pattern:**
```python
# Separate business logic from Office interactions
class ExcelProcessor:
    def __init__(self, excel_app):
        self.excel_app = excel_app
    
    def process_data(self, data):
        # Business logic - easily testable
        pass
    
    def write_to_excel(self, processed_data):
        # Office interaction - mock in tests
        pass
```

### System Automation Scripts

**Unit Testing:**
- ✅ Mock Windows API calls
- ✅ Test logic flow

**Integration Testing:**
- ✅ Test with real Windows APIs in CI
- ✅ Safe system operations only

### File/Registry Operations

**Unit Testing:**
- ✅ Mock file system operations
- ✅ Test business logic

**Integration Testing:**
- ✅ Test with temporary files/registry keys
- ✅ Cleanup after tests

## CI/CD Strategy

### GitHub Actions Workflow

```yaml
# Automated tests that run in CI
jobs:
  unit-tests:
    runs-on: windows-latest
    # All mocked unit tests
  
  integration-tests:
    runs-on: windows-latest
    # Safe integration tests only
  
  # Manual trigger for Office tests
  office-tests:
    runs-on: self-hosted  # Requires pre-configured environment
    when: manual
```

### Local Development

- Full test suite including Office interactions
- Test data and fixtures provided
- Clear separation between automated and manual tests

## Test Organization

```
tests/
├── unit/                 # Fully automated
│   ├── test_utils/
│   ├── test_core/
│   └── test_mocks/
├── integration/          # Partially automated
│   ├── test_system/
│   └── test_safe_scripts/
├── manual/               # Manual testing
│   ├── test_office/
│   ├── fixtures/
│   └── checklists/
└── conftest.py           # Shared fixtures and mocks
```

## Conclusion

**Automated Testing Coverage:** ~70-80% of codebase
- All utility functions and business logic
- Safe system interactions
- CLI and configuration handling

**Manual Testing Required:** ~20-30% of codebase
- Office application interactions
- Complex system modifications
- End-to-end user workflows

This hybrid approach ensures good test coverage while acknowledging the practical limitations of testing Windows application automation in CI environments.
