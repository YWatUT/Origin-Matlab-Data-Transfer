function SetColUserParam(obj, Col, up_name, up_value)
% set a user parameter for the worksheet
% related commands:
% invoke (h, 'Execute', 'wks.userParam2 = 1')  turns on the second user parameter
% invoke (h, 'Execute', 'wks.userParam2$ = name')
% invoke(h, 'Execute', 'col(2)[D2]$ = aa ')
% get(h, 'UserDefLabel','1') gets the name of the user label


% Behaviour of user parameters in origin:
% When a work sheet is created, no user parameters will be activated. You
% can do invoke(h, 'Execute', 'col(1)[D1]$ = aa') commands but they won't
% actually do/set anything. if you do get(h, 'UserDefLabel','0') command it
% will return an empty char array.

% Once you turn on a user parameter by invoke(h, 'Execute', 'wks.userParam1 = 1'), 
% a UserDefined parameter will appear and stay there. You can change it's
% name by running invoke (h, 'Execute', 'wks.userParam1$ = name'), but even
% if you turn the visibility off by invoke(h, 'Execute', 'wks.userParam1 = 0'),
% the parameter will stay there. If you try to get the name via get(h, 'UserDefLabel','0')
% it will return 'name' even though it's no longer visible in the sheet.
% That is, untill you do 'wks.userParam1 = 1' again.


for i = 0:5
    up_str = get(obj.CurrentSheet, 'UserDefLabel', num2str(i));
    
    if isempty(up_str)
        % this indicates that the current user parameter hasn't been initialized yet
        invoke(obj.CurrentSheet, 'Execute', ['wks.userParam',num2str(i+1),'$=',up_name])
        invoke(obj.CurrentSheet, 'Execute', ['wks.userParam',num2str(i+1),'=1'])
        invoke(obj.CurrentSheet, 'Execute', ['col(',num2str(Col),')[D',num2str(i+1),']$ =', up_value]);
        return;
    elseif strcmp(up_str, up_name)
        % this indicates that the current user parameter is the same as the passed parameter
        % the corresponding user parameter can be directly updated
        invoke(obj.CurrentSheet, 'Execute', ['col(',num2str(Col),')[D',num2str(i+1),']$ =', up_value]);
        return
    else
        % this indicates that the current user parameter exist, but is not the same as the passed
        % parameter, so move on to the next
        continue
    end

end

% If the for loop is exited without return, it means that the passed user parameter is not the same
% as the 6 defined user parameters -- maybe a little bit too much

error('OriginClass.SetColUserParam : trying to go over six user parameters -- this is not yet supported')

end