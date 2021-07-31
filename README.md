# Utility to bulk upload products

The purpose of this utility is to make it easy for end users to rapidly add products to myshop53 store.

## Installation and Pre-requisites

1. Ensure that you have python 3.6 or later installed on your machine
2. Ensure that python dependencies are installed using below command

```
python3 -m pip install -r requirements.txt
```

## Usage

```
$ python3 myshop.py -h
usage: myshop.py [-h] [-c CONFIG] [-p PRODUCTS] [-s] [-e] [-i] [-m]

Utility to bulk create/update products, categories & manufacturers in your myshop store.

optional arguments:
  -h, --help            show this help message and exit
  -c CONFIG, --config CONFIG
                        Store configuration file
  -p PRODUCTS, --products PRODUCTS
                        File that will contain the products informations
  -s, --skip-images     Skip importing image files associated with the products
  -e, --products_export
                        Export products from myshop store to local files
  -i, --products_import
                        Import products to myshop store
  -m, --maintain_state  Maintain import state so that the next import runs from where it stopped
  ```
  
  ## Export Products
  
  Product metadata can be exported in the microsoft excel format, updated or new products added and can be imported back.
  
  ```
  python3 myshop.py -c config.yaml -p products.xlsx -e
  ```
  
  ## Import Products & related images
  
  Product data and images can be uploaded to myshop53 store. 'Images' column can contain one or more comma separated image paths relative to working directory of this utility. The image paths will be looked up in the local directory and uploaded to the store along with product creation/updation. There are multiple options available with import, please see Usage for more details.
  
  ```
  python3 myshop.py -c config.yaml -p products.xlsx -i
  ```
  
  
  
  
  
