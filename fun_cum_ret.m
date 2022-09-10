% Cum_Ret=fun_cum_ret(X,I,FW_BW,PERCENT_LOG)
%
% X is T-by-1 Vec of Returns.
%
% I is a [3,6,12] for example (number of periods to cummulate).
%
% FW_BW_CT is 'FW' or 'BW' or 'CT' - Cum FW or BW or CT (center)
%
% LOG_PERCENT is 'LOG' or 'PERCENT' - LOG is regular sum (for log diffs)
%                             PERCENT is compounded ret (for crsp returns)



function Cum_Ret=fun_cum_ret(X,I,FW_BW,LOG_PERCENT)

Cum_Ret=zeros(length(X),length(I));

if strcmp(FW_BW,'BW') && strcmp(LOG_PERCENT,'LOG')
   for i=1:length(I)
       for t=I(i):length(X)
           for j=1:I(i)
               Cum_Ret(t,i)=Cum_Ret(t,i)+X(t-(j-1));
           end
       end
   end
end

if strcmp(FW_BW,'BW') && strcmp(LOG_PERCENT,'PERCENT')
   for i=1:length(I)
       for t=I(i):length(X)
           for j=1:I(i)
               Cum_Ret(t,i)=(1+Cum_Ret(t,i))*(1+X(t-(j-1)))-1;
           end
       end
   end
end

if strcmp(FW_BW,'FW') && strcmp(LOG_PERCENT,'LOG')
   for i=1:length(I)
       for t=1:length(X)-(I(i)-1)
           for j=1:I(i)
               Cum_Ret(t,i)=Cum_Ret(t,i)+X(t+(j-1));
           end
       end
   end
end

if strcmp(FW_BW,'FW') && strcmp(LOG_PERCENT,'PERCENT')
   for i=1:length(I)
       for t=1:length(X)-(I(i)-1)
           for j=1:I(i)
               Cum_Ret(t,i)=(1+Cum_Ret(t,i))*(1+X(t+(j-1)))-1;
           end
       end
   end
end

if strcmp(FW_BW,'CT') && strcmp(LOG_PERCENT,'LOG')
   for i=1:length(I)
       for t=((I(i)+1)/2):(length(X)-((I(i)-1)/2))
           for j=1:I(i)
               Cum_Ret(t,i)=Cum_Ret(t,i)+X(t-(I(i)-1)/2+j-1);
           end
       end
   end
end

if strcmp(FW_BW,'CT') && strcmp(LOG_PERCENT,'PERCENT')
   for i=1:length(I)
       for t=((I(i)+1)/2):(length(X)-((I(i)-1)/2))
           for j=1:I(i)
               Cum_Ret(t,i)=(1+Cum_Ret(t,i))*(1+X(t-(I(i)-1)/2+j-1))-1;
           end
       end
   end
end

end