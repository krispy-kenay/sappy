# sappy
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
[![Python](https://img.shields.io/badge/python-3.10+-blue)](https://www.python.org/)

Simple interface to access the SAP scripting engine via python with some helper functions

## Usage

### Installation

The intended way to install this package is to download the latest version from [releases](https://github.com/krispy-kenay/vasp-analyzer/releases) and then use pip to install it:

```bash
pip install sappy-<version>.tar.gz
```

This package can also be directly installed from github using `pip`:

```bash
pip install git+https://github.com/krispy-kenay/sappy.git
```

### Import
Import the module like this:

```python
from sappy import Client
```

### Open a session
From here, open a new session using a context manager:

```python
with Client().new_session("Server_Name") as session:
    # Perform your logic here
```

It is highly recommended to use a context manager so that the session will close once it has been used, but you can also open one without (just make sure to close it afterwards):

```python
session = Client().new_session("Server_Name")
# Perform your logic here
session.close()
```

### Queries
Once a session is open, you can do some basic queries. 

| Command | Explanation |
| --- | --- |
| `session.open_transaction("Transaction_Id")` | Open a transaction |
| `session.close_transaction()` | Close the currently open transaction |
| `session.send_key(key)` | Send a key to SAP (ex. 0 for "Enter")|  
| `session.find_elements("Search_for")` | Return all paths to elements that contain this substring |
| `session.find_element("Search_for")` | Return the element with an exact match directly |
| `session.update_field(field_id, "Set_to_this")` | Change the content of a field |
| `session.get_table(table_id)` | Return the text content as a python list of lists |

## Tips & Tricks
It might seem a bit unintuitive at first to get started, since finding the path of an element is not exactly straightforward. But using:

```python
session.find_elements("")
```

Returns a list containing all paths/ids of all the elements in the current session. 

From there it becomes much easier to find what you want. To filter/find what you want, remember that all elements contain a "short" description of what they are in their path/id:

- `tbl...` Is a table that you can access using `get_table`
- `ctxt...` Is an input field that you can set using `update_field`
- `txt...` Is a regular text field, return the element using `find_element` and read it using `element.text`

## Example
This is a short example of how you would open a new session, then open a transaction, set a field (with id `EX1`) to "example text", press enter and finaly download a table from there.

```python
from sappy import Client

with Client().new_session("Example Server") as session:
    session.open_transaction("EX1")
    session.update_field("ctxtINPT","example text")
    session.send_key(0)
    table = session.get_table("usr/tblSAPEX")
```