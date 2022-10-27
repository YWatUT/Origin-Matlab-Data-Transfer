% Imagine a use senario
WorksheetName = '[test]Sheet2';
data = (1:10)';

Origin = OriginObjClass;                             % Initializing handle to origin session

Origin.Send(WorksheetName, data, 1);      % Send data to worksheet 'WorksheetName', starting from (row 1, ignored by default) column 1
Origin.SetCol(1, 'Name', 'Xlabel');       % Set the name of the current column, which is the one just sent into Origin for data, to 'Xlabel'

Origin.Send(WorksheetName, data, 2);                   % Send data to worksheet 'WorksheetName', starting from row 5, column 2;
Origin.SetCol(2, 'Name', 'Ylabel', 'Unit', 'Y Unit');       % You can also set multiple parameters in one command. 
Origin.SetColUserParam(2, 'Temperature', '22');
                                                         % Supported are 'Name', 'Unit', and I think 'Comment'
Origin.Send(WorksheetName, data, 3);
Origin.SetCol(3, 'name', 'zlabel', 'type', 'z');
Origin.SetColUserParam(3, 'Temperature', '22');
Origin.SetColUserParam(3, 'UP2', 'anything');
Origin.Release;                                             % Release handle so the Origin session can be properly closed