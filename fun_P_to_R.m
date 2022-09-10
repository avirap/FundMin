function Y=fun_P_to_R(X,LOG_PERCENT)

% Y=fun_P_to_R(X,LOG_PERCENT)
%                 X=[P_1 P_2 ...]
%                 Y=[R_1 R_2 ...]
% First Row of Y is [NaN NaN ...]
% LOG_PERCENT is either Log-Difference ('LOG')
%                    or P/P(-1)-1      ('PERCENT')
Y=zeros(size(X));
if strcmp(LOG_PERCENT,'LOG')
    for i=1:size(X,2)
        Y(2:end,i)=log(X(2:end  ,i))...
                  -log(X(1:end-1,i));                           
    end
end

if strcmp(LOG_PERCENT,'PERCENT')
    for i=1:size(X,2)
        Y(2:end,i)=X(2:end  ,i)./...
                     X(1:end-1,i)-ones(size(X,1)-1,1);
    end
end

Y(1,:)=NaN;

end