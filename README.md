# rcrypt

The rcrypt package provides easy symmetric file encryption using GPG with cryptographically strong defaults. Only symmetric encryption is supported.

## Examples

To encrypt a file 

``` r
encrypt("path/to/your/file.csv")
encrypt("path/to/your/file.csv", output = "path/to/your/file.csv.gpg")
```

To decrypt a file 

``` r
decrypt("path/to/your/file.csv.gpg")
decrypt("path/to/your/file.csv.gpg", output = "path/to/your/file.csv")
```
## Installation

To install from CRAN 

``` r
install.packages("rcrypt")
```

Note: rcrypt depends on GPG. GPG must be installed before using rcrypt.

### Install GPG on Windows
Download and install from <http://www.gpg4win.org/>. Make sure GPG is located in the PATH. You can test this by running `gpg --version` from the command line. 

### Install GPG on Linux
GPG is bundled with most Linux distributions. You can test this by running `gpg --version` from the command line.

### Install GPG on OS X
Download and install from <https://gpgtools.org/>. You can test if it's installed by running `gpg --version` from the command line. 
