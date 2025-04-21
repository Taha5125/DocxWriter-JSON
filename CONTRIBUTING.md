# Contributing to DocxWriter

Thank you for your interest in contributing to DocxWriter! This document provides guidelines and instructions for contributing to this project.

## Code of Conduct

By participating in this project, you agree to abide by our Code of Conduct.

## How to Contribute

1. Fork the repository
2. Create a new branch for your feature or bugfix
3. Make your changes
4. Run the tests to ensure your changes don't break existing functionality
5. Submit a pull request

## Development Setup

1. Clone your fork of the repository
2. Create a virtual environment:

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Running Tests

To run the tests, use the following command:

```bash
python -m unittest tests.py
```

## Pull Request Process

1. Update the README.md with details of changes if needed
2. Update the documentation if you're changing functionality
3. The pull request will be merged once you have the sign-off of at least one maintainer

## Reporting Bugs

If you find a bug, please create an issue with the following information:

- A clear, descriptive title
- Steps to reproduce the issue
- Expected behavior
- Actual behavior
- Screenshots if applicable
- Environment details (OS, Python version, etc.)

## Feature Requests

We welcome feature requests! Please create an issue with the following information:

- A clear, descriptive title
- A detailed description of the feature
- Any examples or mockups if applicable

## Coding Standards

- Follow PEP 8 style guide
- Use type hints for function parameters and return values
- Write docstrings for all functions and classes
- Keep functions focused on a single responsibility
- Write meaningful commit messages

## License

By contributing to DocxWriter, you agree that your contributions will be licensed under the MIT License.
