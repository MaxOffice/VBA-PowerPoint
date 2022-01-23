# VBA-PowerPoint Contribution Guide

## Welcome

Thank you for your interest in VBA Macros for PowerPoint. This guide will help if you want to contribute your own macros to this collection.

We follow a few rules while maintaining this repository. The principal rule is: development happens in the `development` branch. The `main` branch merges directly from `development`, and tags and releases happen from `main`. So, if you want to work on this:

1. Fork this repository.
2. Do your work on the `development` branch.
3. When you are ready, raise a pull request.

We have a few processes and coding guidelines. These, along with some terminology we have used to define them, are listed below. 

## Terminology used in this guide

### Packages

A _package_ provides some unique features, and is represented by a directory under the `packages` directory in this repository. Packages may contain VBA modules, class modules, and userforms.

### Actions

An _action_ is a feature provided by a package, which can be invoked from the `Macros` dialog. Actions are implemented in VBA as `Public Sub` routines, with no parameters.

### APIs

An _api_ is a feature provided by a package, which can be invoked from VBA code. APIs are implemented in VBA as `Public Function` or `Public Sub` routines, with at least one parameter.

### Custom UI

_Custom UI_ is ribbon tabs, buttons etc. added via XML files added to Office documents. Custom UI is implemented using a file called `customUI14.xml` in a subdirectory called `CustomUI` under each package directory. If the custom UI includes buttons with cutom images, the image files also must be stored in the `CustomUI` subdirectory. The images must be stored in .png format.

### UIActions

An _UIAction_ is a feature invoked by the click of a button on the ribbon. UIActions are implemented in VBA as `Public Sub` routines with a single parameter of type `IRibbonControl`.

### "should" and "must"

In the processes and guidelines below, the terms "should" and "must" have the expected meaning. We will send any pull request that violates a "must" guideline back for review, whereas "should" guideline violations may in some cases be accepted as is.


## Processes

### Adding a new package

1. Create a new subdirectory for your package under the `packages` directory. Give it a unique name.
2. Create a subdirectory called `Tests` under the package directory.
3. Create a .PPTM file under `Tests`. This will be your dev/test environment.
4. Add you functionality to the .PPTM file, following the guidelines below.
5. Be sure to include test cases in the .PPTM. Include one case per slide.
6. Once all code has been tested, export your modules (`.bas` files), class modules (`.cls` files) and userforms (`.frm` and `.frx` files) to the package directory.

## Coding Guidelines

### Mandatory guidelines

1. All UserForm and Module members (varibles, constants, `Sub`s, `Function`s) must be `Private` by default. Class modules must have their `Instancing` property set to `1 - Private` by default.
2. A package may provide one or more _actions_. These must be declared as `Public Sub` routines in Modules, with a "PascalCased" name. Action subroutines should ideally handle all errors, and provide appropriate error messages.
3. A package may provide one or more _apis_. These must be declared as `Public Function` or `Public Sub` routines in Modules, with a "PascalCased" name. API subroutines should raise errors if needed, and not handle everything. If an API returns a Class type, or takes a Class type as a parameter, the relevant Class Module must have the `Instancing` property set to `2 - PublicNotCreatable`.
4. All `MsgBox` calls should include a title parameter, identifying the package which raised the message.

### Naming Conventions

1. `Public` functions and subroutines must be named using "PascalCase".
2. All other members should be named using "camelCase".
3. Module, Class Module and Form names should be in "PascalCase". Form names should end with the word `Form`.
4. The primary action subroutine should have the same name as the package directory.
