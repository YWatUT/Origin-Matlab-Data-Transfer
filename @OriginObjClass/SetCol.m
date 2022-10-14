function SetCol(obj, Col, varargin)  
    % Set the info of the selected column of the current worksheet
    
    % First parse varargin into ['Property', 'String'] pairs
    if mod(length(varargin),2)
        error('OriginClass.SetCol : Bad input, property and value must be in pairs!');
    end
    
    property_array = varargin(1:2:end);
    value_array = varargin(2:2:end);
    
    for i = 1:length(property_array)
        if ~isColProperty(property_array{i})
            error(['OriginClass.SetCol : Bad input, property ',property_array{i},' is invalid!']);
        end
    end

    % Select the current column to input column name, units and etc)
    if ~isempty(obj.CurrentSheet)
        for i = 1:length(property_array)
            col_str = ['col$(' , num2str(Col) , ')'];
            prop_str = PropStr(property_array{i});
            value_str = ValueStr(value_array{i}, property_array{i});
            Command = ['wks.', col_str, '.', prop_str, '=',value_str];
            invoke(obj.CurrentSheet, 'Execute', Command);
            % an example: invoke(obj.CurrentSheet, 'Execute','wks.col$(3).lname$=ColName')
            % alternative style: invoke(obj.CurrentSheet, 'Execute', 'Col(2)[L]$ = ColName '); 
        end
    else 
        return
    end
    
end

%% Utility functions

function Bool = isColProperty(string)
      PropertyList = {
          'name';
          'unit';
          'comment'; 
          'comments';
          'type';
      };
      Bool = any(strcmpi(PropertyList,string));
end

function prop_str = PropStr(string)
    PropertyTable = {
        'name',     'lname$';
        'unit',     'unit$';
        'comment',  'comment$';
        'comments', 'comment$';
        'type',     'type';
    };

      
      % https://www.originlab.com/doc/en/LabTalk/ref/Column-Label-Row-Characters
      % https://www.originlab.com/doc/LabTalk/ref/Wks-Col-obj
      
      prop_str = PropertyTable{ strcmpi(PropertyTable(:,1),string),   2};
end

function value_str = ValueStr(input_str, prop_type)
    if strcmp(prop_type,'type')
        type_table = {
            'Y',            '1';
            'disregard',    '2';
            'Y error',      '3';
            'X',            '4';
            'Label',        '5';
            'Z',            '6';
            'X error',      '7';
        };
        find_index = strcmpi(type_table(:,1), input_str);
        
        if ~any(find_index)
            error(['OriginClass.SetCol : Bad input, type ', input_str, ' is invald!']);
        end
        
        value_str = type_table{find_index ,  2};
    else
        value_str = input_str;
    end
end