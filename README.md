[![Gitpod ready-to-code](https://img.shields.io/badge/Gitpod-ready--to--code-blue?logo=gitpod)](https://gitpod.io/#https://github.com/anthunt/aws-exporter)

AWS Exporter Usage
=============

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/f30922fcbdbc4babaf3bc4b483d4b262)](https://www.codacy.com/gh/anthunt/aws-exporter/dashboard?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=anthunt/aws-exporter&amp;utm_campaign=Badge_Grade)
![Maven Package](https://github.com/anthunt/aws-exporter/workflows/Maven%20Package/badge.svg)
![Java CI with Maven](https://github.com/anthunt/aws-exporter/workflows/Java%20CI%20with%20Maven/badge.svg)

AWSExporter extracts configuration information about resources you create in AWS into an Excel file.

:wrench: Installation
-------------

1. Download [aws-exporter-v3.0.4-RELEASE.zip](https://github.com/anthunt/aws-exporter/releases/download/v3.0.4-RELEASE/aws-exporter-v3.0.4-RELEASE.zip) to your local path and run install
2. Unzip file
3. install jar - it will be download dependencies and make structures
```
$ java -jar aws-exporter-v3.0.4-RELEASE-full.jar
```

:file_folder: Program Structure
-------------
<pre><code>
Your path -
          |- aws-exporter-v3.0.4-RELEASE-full.jar
          |- startAWSExporter.bat (created at install)
          |- lib (created at install)
          |- conf (created when running startAWSExport.bat)
               |- env.bat
               |- awsExporter.data
          |- exports (created at export)
</code></pre>

:arrow_forward: Running Program
-------------

```
$ startAWSExporter.bat
```

<p align="center">
  <img src="https://github.com/anthunt/aws-exporter/blob/3.x/running.png?raw=true">
</p>

:pencil: Output Excel File
-------------
Output file will be create in exports directory.
#### Output File Naming Format : 
```
[Access Name][Region Name] AWSExport-[Export time].xlsx
```

:heavy_check_mark: Support AWS Services
-------------

+ VPC
+ Subnet
+ Route Table
+ Internet Gateway
+ Egress Only Internet Gateway
+ NAT Gateway
+ Customer Gateway
+ VPN Gateway
+ VPN Connection
+ Security Group
+ EC2 Instance
+ EBS Volume
+ ElasticLoadBalancer (ALB, NLB, CLB)
+ ElastiCache
+ RDS
+ Key Management Service
+ Certificate Manager
+ Simple Storage Service
+ Direct Connect Connection
+ Direct Connect Location
+ Virtual Gateway
+ Virtual Interface
+ LAG
+ Direct Connect Gateway
+ Directory Service
