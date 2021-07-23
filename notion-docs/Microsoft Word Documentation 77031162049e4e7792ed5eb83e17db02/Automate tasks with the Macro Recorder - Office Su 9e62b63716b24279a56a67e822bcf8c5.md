# Automate tasks with the Macro Recorder - Office Support

[https://support.microsoft.com/en-us/office/automate-tasks-with-the-macro-recorder-974ef220-f716-4e01-b015-3ea70e64937b](https://support.microsoft.com/en-us/office/automate-tasks-with-the-macro-recorder-974ef220-f716-4e01-b015-3ea70e64937b)

To automate a repetitive task, you can record a macro with the Macro Recorder in Microsoft Excel. Imagine you have dates in random formats and you want to apply a single format to all of them. A macro can do that for you. You can record a macro applying the format you want, and then replay the macro whenever needed.

When you record a macro, the macro recorder records all the steps in Visual Basic for Applications (VBA) code. These steps can include typing text or numbers, clicking cells or commands on the ribbon or on menus, formatting cells, rows, or columns, or even importing data from an external source, say, Microsoft Access. Visual Basic Application (VBA) is a subset of the powerful Visual Basic programming language, and is included with most Office applications. Although VBA gives you the ability to automate processes within and between Office applications, it is not necessary to know VBA code or computer programming if the Macro Recorder does what you want.

It is important to know that you when you record a macro, the Macro Recorder captures almost every move you make. So if you make a mistake in your sequence, for example, clicking a button that you did not intend to click, the Macro Recorder will record it. The resolution is to re-record the entire sequence, or modify the VBA code itself. This is why whenever you record something, it's best to record a process with which you're highly familiar. The more smoothly you record a sequence, the more efficiently the macro will run when you play it back.

Macros and VBA tools can be found on the **Developer** tab, which is hidden by default, so the first step is to enable it. For more information, see [Show the Developer tab](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45).

![Automate%20tasks%20with%20the%20Macro%20Recorder%20-%20Office%20Su%209e62b63716b24279a56a67e822bcf8c5/be1200dc-fdf2-48d0-a8e6-b23a1a885339.png](Automate%20tasks%20with%20the%20Macro%20Recorder%20-%20Office%20Su%209e62b63716b24279a56a67e822bcf8c5/be1200dc-fdf2-48d0-a8e6-b23a1a885339.png)

### Record a macro

There are a few helpful things you should know about macros:

- 

    When you record a macro for performing a set of tasks in a range in Excel, the macro will only run on the cells within the range. So if you added an extra row to the range, the macro will not run the process on the new row, but only the cells within the range.

- 

    If you have planned a long process of tasks to record, plan to have smaller relevant macros instead of having one long macro.

- 

    It is not necessary that only tasks in Excel can be recorded in a macro. Your macro process can extend to other Office applications, and any other applications that support Visual Basic Application (VBA). For example, you can record a macro where you first update a table in Excel and then open Outlook to email the table to an email address.

Follow these steps to record a macro.

1.     

    On the **Developer** tab, in the **Code** group, click **Record Macro**.

    - OR-

    Press **Alt**+**T**+**M**+**R** .

    ![Automate%20tasks%20with%20the%20Macro%20Recorder%20-%20Office%20Su%209e62b63716b24279a56a67e822bcf8c5/b0ddaa7d-64a2-450c-84e5-589fe04f57a8.png](Automate%20tasks%20with%20the%20Macro%20Recorder%20-%20Office%20Su%209e62b63716b24279a56a67e822bcf8c5/b0ddaa7d-64a2-450c-84e5-589fe04f57a8.png)

2.   

    In the **Macro name** box, enter a name for the macro. Make the name as descriptive as possible so you can quickly find it if you create more than one macro.

    **Note:** The first character of the macro name must be a letter. Subsequent characters can be letters, numbers, or underscore characters. Spaces cannot be used in a macro name; an underscore character works well as a word separator. If you use a macro name that is also a cell reference, you may get an error message that the macro name is not valid.

3.  

    To assign a keyboard shortcut to run the macro, in the **Shortcut key** box, type any letter (both uppercase or lowercase will work) that you want to use. It is best to use **Ctrl** + **Shift** (uppercase) key combinations, because the macro shortcut key will override any equivalent default Excel shortcut key while the workbook that contains the macro is open. For instance, if you use **Ctrl+Z** (Undo), you will lose the ability to Undo in that Excel instance.

4.   

    In the **Store macro in** list, select where you want to store the macro.

    In general, you’ll save your macro in the **This Workbook** location, but if you want a macro to be available whenever you use Excel, select [Personal Macro Workbook](https://support.microsoft.com/en-us/office/copy-your-macros-to-a-personal-macro-workbook-aa439b90-f836-4381-97f0-6e4c3f5ee566) . When you select **Personal Macro Workbook**, Excel creates a hidden personal macro workbook (Personal.xlsb) if it does not already exist, and saves the macro in this workbook.

5.   

    In the **Description** box, optionally type a brief description of what the macro does.

    Although the description field is optional, it is recommended you enter one. Also, try to enter a meaningful description with any information that may be useful to you or other users who will be running the macro. If you create a lot of macros, the description can help you quickly identify which macro does what, otherwise you might have to guess.

6.  

    Click **OK** to start recording.

7.  

    Perform the actions that you want to record.

8.    

    On the **Developer** tab, in the **Code** group, click **Stop Recording** .

    - OR-

    Press **Alt**+**T**+**M**+**R** .

### Working with recorded macros in Excel

In the **Developer** tab, click **Macros** to view macros associated to a workbook. Or press **Alt+ F8**. This opens the **Macro** dialog box.

![Automate%20tasks%20with%20the%20Macro%20Recorder%20-%20Office%20Su%209e62b63716b24279a56a67e822bcf8c5/15e4508a-6535-4ef2-9b0a-a9a6b2ef9179.png](Automate%20tasks%20with%20the%20Macro%20Recorder%20-%20Office%20Su%209e62b63716b24279a56a67e822bcf8c5/15e4508a-6535-4ef2-9b0a-a9a6b2ef9179.png)

**Caution:** Macros cannot be undone. Before you run a recorded macro for the first time, make sure that you’ve either saved the workbook where you want to run the macro, or better yet work on a copy of the workbook to prevent unwanted changes. If you run a macro and it doesn’t do what you want, you can close the workbook without saving it.

Here's further information on working with macros in Excel.

[Untitled](Automate%20tasks%20with%20the%20Macro%20Recorder%20-%20Office%20Su%209e62b63716b24279a56a67e822bcf8c5/Untitled%20Database%20f21d8b795b6d421f8ab0c06164fd06c7.csv)

### Working with recorded code in the Visual Basic Editor (VBE)

You can work with recorded code in the Visual Basic Editor (VBE) and add your own variables, control structures, and other code that the Macro Recorder cannot record. Since the macro recorder captures almost every step you take while recording, you can also clean up any unnecessary recorded code that serves no purpose in the macro. Reviewing your recorded code is a great way to learn or sharpen your VBA programming skills.

For more information on modifying recorded code in an example, see [Getting Started with VBA in Excel](https://msdn.microsoft.com/en-us/library/office/ee814737(v=office.14).aspx).

## Need more help?

You can always ask an expert in the [Excel Tech Community](https://techcommunity.microsoft.com/t5/Excel/ct-p/Excel_Cat), get support in the [Answers community](https://go.microsoft.com/fwlink/?linkid=827514), or suggest a new feature or improvement on [Excel User Voice](https://excel.uservoice.com/).