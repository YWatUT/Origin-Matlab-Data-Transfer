% ParsePrep = {'1',1,'2',2};
PropFlag = 1;

% Properties = {};
% Values = [];
% 
% while ~isempty(ParsePrep)
%       ReadIn = ParsePrep{1};
%       ParsePrep = ParsePrep(2:end);
%       if PropFlag && ischar(ReadIn)               % Should be reading in a property string
%             Properties = [Properties, ReadIn];
%             PropFlag = 0;
%       elseif ~PropFlag && isnumeric(ReadIn)    % Should be reading in a value for the previous property string
%             Values = [Values, ReadIn];
%             PropFlag = 1;
%       else 
%             disp('Bad input, cannot set column info');
%             return 
%       end
% 
% end