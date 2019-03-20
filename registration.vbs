Option Explicit

Dim answer, closeOne
Dim id, firstName, lastName
Dim fso, oFile
Dim state, firstState
Const WRITE = 2
Const READ = 1
Const APPEND = 8

Set fso = CreateObject("Scripting.FileSystemObject")
Set oFile = fso.OpenTextFile("I:\Production Shared Docs\Production Programs\Engraving Programs\vb_app\tykma_project\operators.txt", APPEND, true)


answer = MsgBox("Do you want to register a new employee?",_
		vbOKCancel+vbSystemModal,_
		"Registration")

if answer=1 then
	state = false
	Do
		id = InputBox("Please enter Enployee's ID number",_
		 "Enployee's ID number",_
		 "Enter the employee's Id")

		if id = "Enter the employee's Id" then
			closeOne = msgbox("Please, enter the employee's ID", 1)
				if closeOne = 2 then
					exit do
				end if
		elseif id = "" then
			exit do
		elseif len(id) <> 3  OR not isNumeric(id) then
			closeOne = msgbox("Please, enter a valid id", 1)
				if closeOne = 2 then
					exit do
				end if
		else
			firstState = false
			Do
			firstName = InputBox("Please, enter the employee's FIRST NAME",_
					     "Employee's FIRST NAME",_
					     "Enter the employee's FIRST NAME")
			if firstName = "" then
				exit do
			elseif firstName = "Enter the employee's FIRST NAME" then
				closeOne = msgbox("Please, enter the FIRST NAME", 1)
				if closeOne = 2 then 
					exit do
				end if
			else
				lastName = InputBox("Please, enter the employee's LAST NAME",_
						    "Employee's LAST NAME",_
						    "Enter the employee's LAST NAME")
			end if
			Loop until firstState = true
			oFile.Write id
			state = true
		end if
	Loop until state = true
end if
