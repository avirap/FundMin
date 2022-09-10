function D_RI=fun_LOAD_D_RI(FILENAME)

[D_RI,D_RI_dates]=xlsread(['\EXCEL\',FILENAME],'Sheet1');                                    % Import
D_RI_dates       =D_RI_dates(2:end,1);                                           % Dates
D_RI_dates       =fun_dates(D_RI_dates,'D','cell','mm/dd/yyyy','datenum','D');   % DateNum Conversion
D_RI             =[D_RI_dates,D_RI];                                         % [Dates,Data]
D_RI(:,2)        =fun_P_to_R(D_RI(:,2),'PERCENT');
D_RI(1,:)        =[];

end