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


% For the starting version, ONLY use the first user parameter! Clean up and
% expand the functionality later. I think in most cases one will be enough

up_str = get(obj.CurrentSheet, 'UserDefLabel', '0');

if isempty(up_str)
    % this indicates that user parameter hasn't been initialized yet
    invoke(obj.CurrentSheet, 'Execute', ['wks.userParam1$=',up_name])
    invoke(obj.CurrentSheet, 'Execute', 'wks.userParam1=1')
elseif ~strcmp(up_str, up_name)
    % if the existing user parameter is not the same as the entering one
    error('OriginClass.SetColUserParam : trying to set a different user parameter -- this is not yet supported')
end

% If this if statement is exited successfully, this means that there is an
% existing user parameter with the same name -- safe to proceed

invoke(obj.CurrentSheet, 'Execute', ['col(',num2str(Col),')[D1]$ =', up_value]);

end