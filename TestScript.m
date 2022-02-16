% Imagine a use senario


Origin = OriginObjClass;                             % Initializing handle to origin session

Origin.Send(WorksheetName, data, 1);      % Send data to worksheet 'WorksheetName', starting from (row 1, ignored by default) column 1
Origin.SetCol(1, 'Name', 'Xlabel');                               % Set the name of the current column, which is the one just sent into Origin for data, to 'Xlabel'

Origin.Send(WorksheetName, data, 5, 2);                   % Send data to worksheet 'WorksheetName', starting from row 5, column 2;
Origin.SetCol(2, 'Name', 'Ylabel', 'Unit', 'Y Unit');       % You can also set multiple parameters in one command. 
                                                                                     % Supported are 'Name', 'Unit', and I think 'Comment', all are case sensitive.

Origin.Release;                                             % Release handle so the Origin session can be properly closed