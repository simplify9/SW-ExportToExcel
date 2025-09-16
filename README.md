
[![GitHub Actions](https://github.com/simplify9/SW-ExportToExcel/actions/workflows/nuget-publish.yml/badge.svg)](https://github.com/simplify9/SW-ExportToExcel/actions/workflows/nuget-publish.yml)
[![NuGet](https://img.shields.io/nuget/v/SimplyWorks.ExportToExcel?style=for-the-badge)](https://www.nuget.org/packages/SimplyWorks.ExportToExcel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](LICENSE)

| **Package**       | **Version** |
| :----------------:|:----------------------:|
| SimplyWorks.ExportToExcel | [![NuGet](https://img.shields.io/nuget/v/SimplyWorks.ExportToExcel?style=for-the-badge)](https://www.nuget.org/packages/SimplyWorks.ExportToExcel/) |



## Introduction 
*ExportToExcel* is a library that provides extensions to [IEnumerable Interface](https://docs.microsoft.com/en-us/dotnet/api/system.collections.ienumerable?view=netcore-3.1) extensions. 

## Getting Started
*ExportToExcel* is available as a package on [NuGet](https://www.nuget.org/packages/SimplyWorks.ExportToExcel/). 

To use *ExportToExcel*, you will require the [`Documentformat.OpenXml`](https://www.nuget.org/packages/DocumentFormat.OpenXml/) package. 

## Functions Available 
1. *ExportToExcel*

ExportToExcel has 3 overloads. You can pass it an enumerable of data, where the column names are generated from TEntity's properties. One where you can pass both an enumerable of data and an enumerable of string, where the column names will be taken from this enumerable. And finally, one where you can pass it a dictionary that will be used to generate the column names.

2. *WriteExcel*

WriteExcel has 6 overloads, taking in column names from a dictionary, as well as some data in an IEnumberable array. It also takes in a file stream to start writing onto that very file with the data it's collected. 

## Getting support 👷
If you encounter any bugs, don't hesitate to submit an [issue](https://github.com/simplify9/ExportToExcel/issues). We'll get back to you promptly!

