# python-finance
Collection of utilities for python-finance


Excel with Python
=================

In finance, Microsoft Excel is used as a handy tool for bond traders and is useful in banking operations, as well as task automations using Visual Basic for Applications (VBA). Excel supports the use of Component Object Model (COM) to extend the functionality for custom tasks. This is achieved with the use of COM add-ins as an in-process COM server. With VBA, a wrapper can be created for the COM add-in function so that the COM component can be integrated as a worksheet cell formula function. COM allows the reuse of objects across different software and hardware environments to interface with each other, without the knowledge of its internal implementation. It allows an object to be created in several languages, such as C, C++, Visual Basic, Delphi, or Python.

In this tutorial, we will learn how to build a COM server in Python. We will then create a COM client in Microsoft Excel and interface with the COM server to perform numerical pricing on the call and put options. We will use the Black-Scholes model, the binomial tree model, and the trinomial lattice model from the earlier tutorials covered in this book for the COM server implementation. By linking to the cell values in Excel, or a market data source subscription within the worksheet cells, we can compute the theoretical option prices on the fly.

In this tutorial, we will cover the following topics:

* Overview of the Component Object Model (COM)
* Understanding Excel for finance and COM
* Prerequisites for building a COM server
* Building the Black-Scholes model COM pricing server
* Registering and unregistering the COM server
* Building the Cox-Ross-Rubinstein binomial tree COM server
* Building the trinomial lattice model COM server
* Setting up VBA functions to build a COM client in Excel
* Setting up parameters in Excel to invoke the COM client-server interface
* Computing the theoretical option prices on the fly in Excel

Overview of COM
===============

COM allows the reuse of objects across different software and hardware environments to interface with each other, without the knowledge of its internal implementation. COM is a proprietary standard and is commonly associated with Microsoft's COM. COM forms the basis for Microsoft's other technologies, including ActiveX, COM+, and Document Component Object Model (DCOM).

COM allows an object to be created in several languages, such as C, C++, Visual Basic, Delphi, or Python. Using COM-aware components, COM classes are built as binary standards. Each COM component has its own class identifier (CLSID), which are globally unique identifiers (GUIDs), used for identification when used on a runtime framework. To locate a COM library, the Microsoft Windows registry is used to list all the available class and interface libraries as GUIDs.


Excel for Finance
=================
The spreadsheet application in the Microsoft Office suite was designed for statistical, engineering, and financial data management. In finance, Microsoft Excel is used as a handy tool for bond traders and an integral part of banking operations to task automations using VBA. For example, built-in Excel functions, such as TBILLYIELD and DURATION, helps you calculate the yield of a T-bill and the Macaulay duration of a bond and displays these values onto a cell.

Excel supports the use of COM to extend the functionality for custom tasks. This is achieved with the use of COM add-ins as an in-process COM server. With VBA, a wrapper can be created for the COM add-in function so that the COM component can be used as a worksheet cell formula function.

In this tutorial, we will take a look at building a COM server in Python. We will then use Microsoft Excel, as our source of data parameters, to perform numerical pricing with the COM object. Using this basic example, we can then extend the functionality of the COM objects for many uses, not limited to real-time trading and pricing.

Building a COM server
=====================
In this section, we are concerned with building the server component of the COM interface. We will first take a look at the prerequisites for building the server components using Python. Then, we will proceed to build an option pricing the COM server using some of the topics we covered in Chapter 4, Numerical Procedures.

Prerequisites
-------------

The COM interface is an industry standard by Microsoft; therefore, the following software is required to complete this tutorial:

* Microsoft Windows XP operating system or later
* Microsoft Excel 2003 or later
* Python 2.7 or later with SciPy and NumPy packages
* The `pythoncom` module

Getting the `pythoncom` Module
------------------------------

The pythoncom module contains Python extensions for Microsoft Windows. The files are available freely as pywin32 on SourceForge at http://sourceforge.net/projects/pywin32/files/. To download the executable file, navigate to the pywin32 folder, and select the latest available build. Download the installer executable that is compatible with your system. Note that there is one download package for each supported version of Python. Be sure to check for the version of Python installed in your environment and download the corresponding package. Some packages have a 32-bit and a 64-bit version available. You must download the one that corresponds to the Python you have installed. Even if you have a 64-bit computer, if you installed a 32-bit version of Python, you must install the 32-bit version of pywin32.

Once the executable file is downloaded to your hard drive, run the installer, and follow the onscreen instructions to add the pythoncom module to your Python environment.

Building the Black-Scholes Model COM server
-------------------------------------------

Let's build a simple COM server using the classic Black-Scholes options pricing model to calculate the theoretical value of a call or a put option. The calculator is implemented as the BlackScholes class with a method named pricer that accepts the current underlying price, strike price, annualized interest rate, time left to maturity in terms of years, volatility of the underlying instrument, and annualized dividend yield as its input parameters. The full COM server code in Python is given as follows:

```python

""" Black-Scholes pricer COM server """
import numpy as np
import scipy.stats as stats
import pythoncom

class BlackScholes:
    _public_methods_ = ["call_pricer", "put_pricer"]
    _reg_progid_ = "BlackScholes.Pricer"
    _reg_clsid_ =  pythoncom.CreateGuid()

    def d1(self, S0, K, r, T, sigma, div):
        return (np.log(S0/K) + ((r-div) + sigma**2 / 2) * T)/ \
               (sigma * np.sqrt(T))

    def d2(self, S0, K, r, T, sigma, div):
        return (np.log(S0 / K) + ((r-div) - sigma**2 / 2) * T) / \
               (sigma * np.sqrt(T))

    def call_pricer(self, S0, K, r, T, sigma, div):
        d1 = self.d1(S0, K, r, T, sigma, div)
        d2 = self.d2(S0, K, r, T, sigma, div)
        return S0 * np.exp(-div * T) * stats.norm.cdf(d1) \
               - K * np.exp(-r * T) * stats.norm.cdf(d2)

    def put_pricer(self, S0, K, r, T, sigma, div):
        d1 = self.d1(S0, K, r, T, sigma, div)
        d2 = self.d2(S0, K, r, T, sigma, div)
        return K * np.exp(-r * T) * stats.norm.cdf(-d2) \
               - S0 * np.exp(-div * T) *stats.norm.cdf(-d1)

if __name__ == "__main__":
    # Run "python binomial_tree_am.py"
    #   to register the COM server.
    # Run "python binomial_tree_am.py --unregister"
    #   to unregister it.
    print "Registering COM server..."
    import win32com.server.register
    win32com.server.register.UseCommandLine(BlackScholes)

```


Note the use of the three magic variables: _public_methods_, _reg_progid_, and _reg_clsid_ in the COM server object. The _public_methods_ variable defines the methods that are exposed to the COM clients. The _reg_progid_ variable defines the name of the COM server that is called from the COM client. The _reg_clsid_ variable contains the unique class identifier in the registry.

Registering and unregistering the COM server
--------------------------------------------

Assuming that the code is saved in the black_scholes.py file, we can compile the COM server and register with the registry:

```
$ python black_scholes.py
Registering COM server…
Registered: BlackScholes.Pricer
```

The COM server is now accessible for COM communications.

To unregister the COM server, the additional --unregister parameter is used:

```
$ python black_scholes.py --unregister
Registering COM server…
Unregistered: BlackScholes.Pricer
```

The COM server is now unregistered and cannot be accessed by the COM clients.

Building the Cox-Ross-Rubinstein binomial tree model COM server
---------------------------------------------------------------

In a previous tutorial, Numerical Procedures, we looked at several options pricing models. One such model is the Cox-Ross-Rubinstein (CRR) model using a binomial tree. Before we can create a second COM server based on this model, let's copy and paste these class files created earlier, namely, BinomialCRROption.py, BinomialTreeOption.py, and StockOption.py, to our working directory.

Now, let's create our COM server using the BinomialCRRCOMServer class and save it as `binomial_crr_com.py`:

```python
""" Binomial CRR tree COM server """
from BinomialCRROption import BinomialCRROption
import pythoncom

class BinomialCRRCOMServer:
    _public_methods_ = [ 'pricer']
    _reg_progid_ = "BinomialCRRCOMServer.Pricer"
    _reg_clsid_ = pythoncom.CreateGuid()

    def pricer(self, S0, K, r, T, N, sigma,
               is_call=True, div=0., is_eu=False):
        model = BinomialCRROption(S0, K, r, T, N,
                                  {"sigma": sigma,
                                   "div": div,
                                   "is_call": is_call,
                                   "is_eu": is_eu})
        return model.price()

if __name__ == "__main__":
    print "Registering COM server..."
    import win32com.server.register
    win32com.server.register.UseCommandLine(BinomialCRRCOMServer)
```

Similar to our Black-Scholes COM server, here the pricer method creates an instance of the BinomialCRROption class and returns the calculated price from the CRR binomial tree model.

Building the trinomial lattice model COM server
In Chapter 4, Numerical Procedures, we also explored the use of a trinomial lattice in options pricing. Let's use this model as our third COM server. Let's copy and paste the related class files, namely, TrinomialLattice.py and TrinomialTreeOption.py, to our working directory.

Create our COM server with the TrinomialLatticeCOMServer class and save it as `trinomial_lattice_com.py`:

```python

""" Trinomial Lattice COM server """
from TrinomialLattice import TrinomialLattice
import pythoncom

class TrinomialLatticeCOMServer:
    _public_methods_ = ['pricer']
    _reg_progid_ = "TrinomialLatticeCOMServer.Pricer"
    _reg_clsid_ = pythoncom.CreateGuid()

    def pricer(self, S0, K, r, T, N, sigma,
               is_call=True, div=0., is_eu=False):
        model = TrinomialLattice(S0, K, r, T, N,
                                 {"sigma": sigma,
                                  "div": div,
                                  "is_call": is_call,
                                  "is_eu": is_eu})
        return model.price()

if __name__ == "__main__":
    print "Registering COM server..."
    import win32com.server.register
    win32com.server.register.UseCommandLine(TrinomialLatticeCOMServer) 
	
```

Now, let's build and register our three COM server Python files with the registry:

```
$ python black_scholes.py
Registering COM server…
Registered: BlackScholes.Pricer

$ python binomial_crr_com.py
Registering COM server…
Registered: BinomialCRRCOMServer.Pricer

$ python trinomial_lattice_com.py
Registering COM server…
Registered: TrinomialLatticeCOMServer.Pricer
```

With our COM server components successfully registered with the registry, we can now proceed to create our COM client in Excel in the next section.

Building the COM client in Excel
================================

In the worksheet cells of Microsoft Excel, we can input a number of parameters for a particular option and numerically compute the theoretical option prices using the COM server components we just built in the earlier section. These functions can be made available in the formula cell using Visual Basic. To begin creating these functions, open the Visual Basic Editor from Excel by pressing the Alt + F11 keys on your keyboard.

Setting up the VBA code
-----------------------

In the Project-VBAProject toolbar window, right-click on VBAProject, select Insert, and click on Module to insert a new module in the Excel workbook:

In the code editor area, paste the following VBA code:

```
Function BlackScholesOptionPrice( _
     ByVal S0 As Integer, _
     ByVal K As Integer, _
     ByVal r As Double, _
     ByVal T As Double, _
     ByVal sigma As Double, _
     ByVal dividend As Double, _
     ByVal isCall As Boolean)
     Set BlackScholes = CreateObject("BlackScholes.Pricer")
     If isCall = True Then
         answer = BlackScholes.call_pricer(S0, K, r, T, sigma, \ dividend)
     Else
         answer = BlackScholes.put_pricer(S0, K, r, T, sigma, \ dividend)
     End If
     BlackScholesOptionPrice = answer
 End Function
```

This will create the COM client component of the Black-Scholes model. The BlackScholesOptionPrice VBA function takes in a number of input parameters from Excel, which we will define later. The CreateObject function is then called and takes the BlackScholes.Pricer input string, which is effectively the name, as defined in the _reg_progid_ variable of the corresponding COM server component. In the COM server, we exposed two methods, call_pricer and put_pricer, to compute and return the Black-Scholes call and put option prices respectively. The selection of this option is determined by the isCall variable, which is true for a call option and false for a put option.

In the same fashion, we can create the COM client functions for our two other pricing methods using the following VBA code:

```
Function BinomialTreeCRROptionPrice( _
    ByVal S0 As Integer, _
    ByVal K As Integer, _
    ByVal r As Double, _
    ByVal T As Double, _
    ByVal N As Integer, _
    ByVal sigma As Double, _
    ByVal isCall As Boolean, _
    ByVal dividend As Double)
    Set BinCRRTree = CreateObject("BinomialCRRCOMServer.Pricer")
    answer = BinCRRTree.pricer(S0, K, r, T, N, sigma, isCall, _
        dividend, True)
    BinomialTreeCRROptionPrice = answer
End Function 
Function TrinomialLatticeOptionPrice( _
    ByVal S0 As Integer, _
    ByVal K As Integer, _
    ByVal r As Double, _
    ByVal T As Double, _
    ByVal N As Integer, _
    ByVal sigma As Double, _
    ByVal isCall As Boolean, _
    ByVal dividend As Double)
    Set TrinomialLattice = _
        CreateObject("TrinomialLatticeCOMServer.Pricer")
    answer = TrinomialLattice.pricer(S0, K, r, T, N, sigma, _
        isCall, dividend, True)
    TrinomialLatticeOptionPrice = answer
End Function
```

Here, the BinomialTreeCRROptionPrice and TrinomialLatticeOptionPrice VBA functions are defined. Similar to the BlackScholesOptionPrice function, the CreateObject function takes in the string value of BinomialCRRCOMServer.Pricer and TrinomialLatticeCOMServer.Pricer, as defined in the _reg_progid_ variable in its respective COM server.

We can compile the code by selecting Debug from the toolbar menu and clicking on Compile VBAProject:

When the code has been successfully compiled, close the Visual Basic Editor window and return to Excel to input our parameters.

Setting up the cells
====================

Let's assume that we would like to price an option with a strike price of 50. The current underlying price is 50 with a volatility of 0.5 and does not pay dividends. The risk-free rate is 0.05 and the time to maturity is 6 months. We will start with a two-step binomial tree and trinomial lattice with N=2.

In Excel, set up the following cells and values:



We are now ready to price our option using dynamic numerical pricing with COM.

Notice that in cells B12 to B14, we are calling the functions that we have defined in the VBA editor. The input values are derived from the values of cells B2 to B8. The Boolean value in B11 determines whether we are pricing a call option or a put option when calling the COM server. Since we are pricing the call options in column B, let's add another column, C to price the put options:

The formulas are the same, as in the previous table, with the exception of the isCall cell reference to C11 instead of B11. This allows us to price a put option.

Our Excel spreadsheet should look something like this:

In a new row, set up the following cells and values:

The call option prices, as computed by the Black-Scholes model, the binomial tree with CRR parameters, and the trinomial lattice model, are 7.5636, 6.7734, and 7.1468 respectively. Likewise, the put option prices are 6.3291, 5.8685, and 6.0823 respectively.

What happens when we change the value of N to a bigger value?


We can see that the values of the binomial tree by the CRR model and the trinomial lattice model converge to the values by the Black-Scholes model as the number of tree step increases.


What else can I do with COM?
============================

On changing the values of N, we can see that the values from our custom-defined functions changes on the fly. This makes it possible for dynamic computations of securities, or even real-time numerical pricing, when connected to a market data feed, where values such as S0 or K are changing every second.

The COM server components are separated from each other. Using Python, we can change the implementation of the COM server using the Python modules, such as NumPy or SciPy, to achieve certain aspects of numerical pricing without relying too much on Excel's built-in functions. This also means that we can interchange and interface components that are not related to Excel. The COM model simply provides a transparent bridge between these components and Excel.


Summary
=======

In this chapter, we looked at the use of the Component Object Model (COM) to allow the reuse of objects across different software and hardware environments to interface with each other, without the knowledge of its internal implementation.

To build the server component of the COM interface, we used the pythoncom module to create a Black-Scholes pricing COM server with the three magic variables: _public_methods_, _reg_progid_, and _reg_clsid_. Using topics in Chapter 4, Numerical Procedures, we created COM server components using the binomial tree by the CRR model and trinomial lattice model. We learned how to register and unregister these COM server components with the Windows registry.

In Microsoft Excel, we can input a number of parameters for a particular option and numerically compute the theoretical option prices using the COM server components we built. These functions are made available in the formula cells using Visual Basic. We created the Black-Scholes model, binomial tree CRR model, and trinomial lattice model COM client VBA functions. These functions accept the same input values from the spreadsheet cells to perform numerical pricing on the COM server. We also saw how to update the input parameters in the spreadsheet that dynamically update the option prices on the fly.
