# Excel Exporter and Importer

## Overview

"Excel Exporter and Importer" is an utility project, written in Java 21 (Maven), which provides utilities for exporting data from a database to Excel files and importing data from Excel files into a database.
This project is intended to be used as a dependency for the "RESTful Export Import API" Spring Boot project.

## Features

- **Excel Exporter:**
  - Export data, from a database, to Excel with customizable headers, and with the ability of making SUM and AVERAGE operations on numeric fields.
  - Support for localization of headers based on user-defined locales in the "RESTful Export Import API".
  - Currently supported locales: en (English), es (Spanish), ja (Japanese), tr (Turkish)

- **Excel Importer:**
  - Import data from Excel files (provided with specified header structure) into a database.
  - Includes validation checks and support for various Excel file structures.

## Installation

### Prerequisites

- Java 21
- Maven

### Building the Project

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/IsaGeriler/excel-exporter-importer.git
   cd excel-exporter-importer
   ```

2. **Build the Project:**

   Maven:
   ```bash
   mvn clean install
   ```

## Usage

### Inside "RESTful Export Import API" Project, add the dependency to your `pom.xml` file:

**Maven:**
```xml
<!-- excel-exporter-importer.jar Dependency -->
    <dependency>
        <groupId>org.example</groupId>
        <artifactId>excel-exporter-importer</artifactId>
	<version>1.0-SNAPSHOT</version>
    </dependency>
```

### Configuration

- **Locale Configuration:**
  - Define supported locales in the "RESTful Export Import API", by creating `Bundle.properties` file for the desired language.

- **Database Configuration:**
  - Configure your database settings in the `application.properties` file of the consuming "RESTful Export Import API" project.

## API Integration

"Excel Exporter and Importer" does not provide its own API but is designed to be used by "RESTful Export Import API", to handle Excel file operations.

## Contributing

1. **Fork the Repository**
2. **Create a New Branch:**
   ```bash
   git checkout -b feature/your-feature
   ```
3. **Commit Your Changes:**
   ```bash
   git commit -am 'Add new feature'
   ```
4. **Push to the Branch:**
   ```bash
   git push origin feature/your-feature
   ```
5. **Create a Pull Request**

## Contact

For any inquiries, please contact [gerilerisaberk@gmail.com](mailto:gerilerisaberk@gmail.com).
