# TaxCalc Web Service

## Overview

TaxCalc is a comprehensive tax computation web service that provides real-time tax calculations for e-commerce and business applications. The service integrates with Taxware's tax computation engine and VeraZip geocoding system to provide accurate tax calculations based on location, product types, and business rules.

## Features

- **Real-time Tax Calculation**: Calculate sales tax, use tax, and other applicable taxes
- **Geocoding Integration**: Uses VeraZip for accurate address geocoding and tax jurisdiction determination
- **Multi-database Support**: Works with multiple database sources (reports, siebeldb)
- **Order Management Integration**: Seamlessly integrates with order management systems
- **Audit Trail**: Comprehensive logging and audit capabilities
- **Flexible Configuration**: Configurable tax rules and business logic
- **Web Service API**: SOAP-based web service for easy integration

## Architecture

### Components

1. **Service.asmx** - Main web service endpoint
2. **Service.vb** - Core business logic and tax computation engine
3. **verazip.ashx** - VeraZip geocoding service handler
4. **web.config** - Configuration and connection settings

### Dependencies

- **Taxware Tax Computation Engine** - Core tax calculation library
- **VeraZip Geocoding System** - Address validation and geocoding
- **SQL Server** - Database backend
- **log4net** - Logging framework
- **.NET Framework 4.0** - Runtime environment

## Installation

### Prerequisites

1. **Windows Server** with IIS 7.0 or higher
2. **.NET Framework 4.0** or higher
3. **SQL Server** 2008 or higher
4. **Taxware Tax Computation Engine** installed and licensed
5. **VeraZip Geocoding System** installed and configured

### Database Setup

1. Run the provided `TaxCalc_Database_Schema.sql` script to create the required database tables
2. Update connection strings in `web.config` with your database server details
3. Configure the following databases:
   - `reports` - Main order and transaction data
   - `siebeldb` - Organization and configuration data

### Configuration

1. **Update web.config**:
   ```xml
   <connectionStrings>
     <add name="hcidb" connectionString="server=YOUR_SERVER;database=YOUR_DATABASE;..." />
     <add name="siebeldb" connectionString="server=YOUR_SERVER;database=siebeldb;..." />
   </connectionStrings>
   ```

2. **Configure Taxware Paths**:
   - Update DLL paths in `Service.vb` to point to your Taxware installation
   - Ensure VeraZip DLLs are accessible from the web service

3. **Set Logging Configuration**:
   - Configure log4net settings in `web.config`
   - Set appropriate log file paths and levels

## API Reference

### Main Tax Calculation Service

**Endpoint**: `Service.asmx`

#### CalculateTax Method

Calculates tax for a specific order and optionally updates the database.

**Parameters**:
- `OrderID` (string) - The ROW_ID of the S_ORDER record
- `Source` (string) - The database where the S_ORDER table is stored
- `UpdateFlg` (string) - If set to "Y", updates the order record with tax amounts
- `AuditFlg` (string) - If set to "Y", creates audit records
- `debug` (string) - If set to "Y", enables debug logging

**Returns**: XML document containing tax calculation results

**Example Request**:
```xml
<soap:Envelope>
  <soap:Body>
    <CalculateTax>
      <OrderID>1-86Z75</OrderID>
      <Source>reports</Source>
      <UpdateFlg>Y</UpdateFlg>
      <AuditFlg>N</AuditFlg>
      <debug>N</debug>
    </CalculateTax>
  </soap:Body>
</soap:Envelope>
```

**Example Response**:
```xml
<tax>
  <order_id>1-86Z75</order_id>
  <total_tax_amount>87.50</total_tax_amount>
  <total_tax_rate>0.0875</total_tax_rate>
  <freight_tax_amount>5.25</freight_tax_amount>
  <freight_tax_rate>0.0875</freight_tax_rate>
  <state>IA</state>
  <city>Cedar Falls</city>
  <zipcode>50613</zipcode>
  <geo_code>19013</geo_code>
  <items>
    <item>
      <row_id>1-ITEM1</row_id>
      <tax_amount>42.00</tax_amount>
      <tax_rate>0.0875</tax_rate>
    </item>
  </items>
</tax>
```

### Geocoding Service

**Endpoint**: `verazip.ashx`

#### Parameters
- `State` (string) - Two-letter state code
- `Zip` (string) - ZIP code (5 or 9 digits)
- `City` (string) - City name

**Returns**: XML document with geocoding results

**Example Request**:
```
GET /verazip.ashx?State=VA&Zip=22209&City=Arlington
```

**Example Response**:
```xml
<results>
  <StateCode>VA</StateCode>
  <Zip1>22209</Zip1>
  <Zip2>22209</Zip2>
  <Geo>51013</Geo>
  <CityName>ARLINGTON</CityName>
  <CntyCode>013</CntyCode>
  <CntyName>ARLINGTON</CntyName>
  <OutsideCity>0</OutsideCity>
</results>
```

## Database Schema

### Core Tables

#### S_ORDER
Main order table containing order information and tax amounts.

**Key Fields**:
- `ROW_ID` - Primary key
- `DISCNT_AMT` - Order discount amount
- `FRGHT_AMT` - Freight amount
- `X_TAX_AMT` - Calculated tax amount
- `X_TAX_RATE` - Tax rate applied
- `X_FRGHT_TAX_AMT` - Freight tax amount
- `X_FRGHT_TAX_RATE` - Freight tax rate

#### S_ORDER_ITEM
Order line items with individual tax calculations.

**Key Fields**:
- `ROW_ID` - Primary key
- `ORDER_ID` - Foreign key to S_ORDER
- `UNIT_PRI` - Unit price
- `QTY_REQ` - Quantity
- `X_TAX_AMT` - Item tax amount
- `X_TAX_RATE` - Item tax rate

#### S_ORG_EXT
Organization/account information.

#### S_ADDR_ORG / S_ADDR_PER
Address tables for organization and personal addresses.

### Configuration Tables

#### S_LST_OF_VAL
Configuration values for taxware settings.

#### TAXWARE_CNTY_FIPS
County FIPS codes for tax jurisdiction determination.

## Business Logic

### Tax Calculation Process

1. **Order Retrieval**: Fetch order and line item data from database
2. **Address Resolution**: Determine ship-to address and geocode
3. **Geocoding**: Use VeraZip to get accurate tax jurisdiction
4. **Tax Computation**: Call Taxware engine with order details
5. **Database Update**: Update order and line items with tax amounts
6. **Audit Logging**: Create audit records if requested

### Tax Rules

- **Product Taxability**: Determined by product codes and configuration
- **Location-based Rates**: Tax rates vary by state, county, and city
- **Freight Taxability**: Configurable freight tax rules
- **Registration Status**: Different rates for registered vs. non-registered customers

## Logging and Monitoring

### Log Levels

- **EventLog**: General application events and errors
- **DebugLog**: Detailed debugging information
- **GZDebugLog**: VeraZip-specific debugging

### Log Configuration

Logging is configured in `web.config` using log4net:

```xml
<log4net>
  <appender name="LogFileAppender" type="log4net.Appender.RollingFileAppender">
    <file value="C:\Logs\TaxCalc.log"/>
    <rollingStyle value="Size"/>
    <maxSizeRollBackups value="10"/>
    <maximumFileSize value="10000KB"/>
  </appender>
</log4net>
```

## Security Considerations

### Data Protection

- **Encryption**: TripleDES encryption for sensitive data
- **Connection Security**: Use encrypted database connections
- **Input Validation**: All inputs are validated and sanitized

### Access Control

- **Authentication**: Configure appropriate authentication methods
- **Authorization**: Implement role-based access control
- **Audit Trail**: Comprehensive logging of all operations

## Troubleshooting

### Common Issues

1. **Taxware DLL Not Found**
   - Verify Taxware installation path
   - Check DLL file permissions
   - Ensure correct architecture (x86/x64)

2. **Database Connection Errors**
   - Verify connection strings in web.config
   - Check database server accessibility
   - Validate user permissions

3. **Geocoding Failures**
   - Verify VeraZip installation
   - Check address data quality
   - Review geocoding service logs

### Debug Mode

Enable debug mode by setting `TaxCalc_debug=Y` in web.config. This will:
- Log detailed operation information
- Show SQL queries being executed
- Display tax calculation parameters

## Performance Optimization

### Database Optimization

- **Indexes**: Ensure proper indexes on frequently queried columns
- **Connection Pooling**: Configure appropriate connection pool sizes
- **Query Optimization**: Use stored procedures for complex queries

### Caching

- **Tax Rate Caching**: Cache frequently used tax rates
- **Geocoding Results**: Cache geocoding results for repeated addresses
- **Configuration Caching**: Cache configuration values

## Maintenance

### Regular Tasks

1. **Log Rotation**: Monitor and rotate log files
2. **Database Maintenance**: Regular index maintenance and statistics updates
3. **Tax Rate Updates**: Update tax rates as needed
4. **Security Updates**: Keep dependencies updated

### Monitoring

- **Performance Metrics**: Monitor response times and throughput
- **Error Rates**: Track and analyze error patterns
- **Resource Usage**: Monitor CPU, memory, and disk usage

## Support and Documentation

### Additional Resources

- **Taxware Documentation**: Refer to Taxware documentation for engine-specific details
- **VeraZip Documentation**: Consult VeraZip documentation for geocoding features
- **SQL Server Documentation**: For database-specific configuration

### Version Information

- **Framework**: .NET Framework 4.0
- **Database**: SQL Server 2008+
- **Web Server**: IIS 7.0+
- **Dependencies**: Taxware, VeraZip, log4net

## License and Legal

This software integrates with third-party components:
- **Taxware**: Commercial tax computation engine
- **VeraZip**: Commercial geocoding system

Ensure proper licensing for all components before deployment.

## Contributing

When contributing to this project:
1. Follow existing code patterns and conventions
2. Add appropriate logging and error handling
3. Update documentation for new features
4. Test thoroughly before submitting changes

## Changelog

### Version 1.0
- Initial release
- Basic tax calculation functionality
- VeraZip geocoding integration
- Database integration
- Comprehensive logging
