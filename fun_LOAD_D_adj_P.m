function D_RI=fun_LOAD_D_adj_P(FILENAME)

[D_RI,D_RI_dates]=xlsread(['\data\',FILENAME]);                                    % Import
D_RI=D_RI(:,5);
D_RI_dates       =D_RI_dates(2:end,1);                                           % Dates
D_RI_dates       =fun_dates(D_RI_dates,'D','cell','yyyy-mm-dd','datenum','D');   % DateNum Conversion
D_RI             =[D_RI_dates,D_RI];                                         % [Dates,Data]
% D_RI=flipud(D_RI);
D_RI(:,2)        =fun_P_to_R(D_RI(:,2),'PERCENT');
D_RI(1,:)        =[];

end