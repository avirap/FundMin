clear all
close all
clc
delete AAA_WorkSpace.mat
%% AddPath
addpath(genpath(('reports')))
addpath(genpath(('data')))

%% BROKER (CASH) ACCOUNTS (BA)
[POOL_local_BA,POOL_text_BA]=xlsread('\data\POOL_BrokerAccounts.xlsx');
% POOL_local_BA=importdata('\data\POOL_BrokerAccounts.xlsx');

POOL_local_BA  =POOL_local_BA(4:end,:);
POOL_ccy_BA    =POOL_text_BA(1,2:3:end);
POOL_names_BA  =POOL_text_BA(2,2:3:end);
POOL_types_BA  =POOL_text_BA(3,2:3:end);
POOL_IDs_BA    =POOL_text_BA(4,2:3:end);
POOL_title_BA  =POOL_text_BA(5,2:3:end);

DATES           =fun_dates(POOL_text_BA(7:end,1),'D','cell','dd/mm/yyyy','datenum','~');
DATES_datevec   =datevec(DATES);
DATES_yyyymmmdd =datestr(DATES(:,1),'yyyy-mmm-dd');
DATES           =fun_last_day_of_month(DATES);
EoM=find(DATES(:,2)==1);
EoY=find(DATES(:,3)==1);
T=length(DATES);
M=length(EoM);
Y=length(EoY);

BA=size(POOL_names_BA,2);                    % NUMBER of BROKER ACCOUNTS
POOL_local_BA=reshape(POOL_local_BA,T,3,BA); % ReShape BROKER ACCOUNTS
clear POOL_text_BA

%% REPORTING PERIOD

BEG=find(DATES==datenum('01-AUG-2022','dd-mmm-yyyy'));
END=find(DATES==datenum('31-AUG-2022','dd-mmm-yyyy'));

%% ACCRUAL (artificial) ACCOUNTS (AA)
[POOL_local_AA,POOL_text_AA]=xlsread('\data\POOL_AccrualAccounts.xlsx');

POOL_local_AA     =POOL_local_AA(:,:);
POOL_ccy_AA       =POOL_text_AA(1,2:5:end);
POOL_names_AA     =POOL_text_AA(2,2:5:end);
POOL_types_AA     =POOL_text_AA(3,2:5:end);
POOL_notes_AA     =POOL_text_AA(4,2:5:end);
POOL_CGIswitch_AA =POOL_text_AA(5,2:5:end);  % CapGn.Income          (CGI) switcher
POOL_MEMswitch_AA =POOL_text_AA(6,2:5:end);  % Distributes.Remembers (MEM) switcher
POOL_IDs_AA       =POOL_text_AA(7,2:5:end);
POOL_title_AA     =POOL_text_AA(8,2:5:end);

AA=size(POOL_names_AA,2);                    % NUMBER of ACCOUNTS
POOL_local_AA=reshape(POOL_local_AA,T,5,AA); % ReShape
clear POOL_text_AA

%% EXPENSE ACCOUNTS (EA) COLLECTIVE/ADMINISTRAION COSTS. PAID OUT OF THE POOL or out of AviRap personal Credit Card. NEEDS to be in USD.
[POOL_dollars_EA,POOL_text_EA]=xlsread('\data\POOL_ExpenseAccounts_usd.xlsx');
 POOL_dollars_EA       =POOL_dollars_EA(:,:);
   POOL_names_EA       =POOL_text_EA(1,2:4:end);
   POOL_types_EA       =POOL_text_EA(2,2:4:end);
   POOL_title_EA       =POOL_text_EA(3,2:4:end);
              EA=size(POOL_names_EA,2); 
     POOL_ccy_EA       =cell(1,EA);
     POOL_ccy_EA(:)    ={'USD'};

                                        % NUMBER of ACCOUNTS
POOL_dollars_EA=reshape(POOL_dollars_EA,T,4,EA);                  % ReShape
for ea=1:EA; POOL_dollars_EA(:,6,ea)=POOL_dollars_EA(:,4,ea); end % move END to COL 6
for ea=1:EA; POOL_dollars_EA(:,4,ea)=POOL_dollars_EA(:,3,ea); end % CAPITALIZATION (as hits EA)
for ea=1:EA; POOL_dollars_EA(:,5,ea)=0;                       end % EFFECT of SUB/RED (agg) is ZERO
clear POOL_text_EA

%% FX
[FX,FX_text]=xlsread('\data\FX.xlsx');
 FX       =FX(:,:);
 FX_ccy   =FX_text(1,2:end);

CC=size(FX_ccy,2); % NUMBER of CCYs
FX_isnan=isnan(FX); FX(FX_isnan)=0; % Replace NaN with zeros.
clear FX_text FX_isnan

%% BROKER ACCOUNTS: CONVERT to USD
for ba=1:BA
    for c=1:CC
        CONVERT_match_TF = strcmp(POOL_ccy_BA(ba),FX_ccy(c));
     if CONVERT_match_TF==1; col=c; end
    end
    POOL_dollars_BA( : ,4,ba)=FX(:,col).*POOL_local_BA(:,3,ba);                          % END of DAY BALANCE
    POOL_dollars_BA( : ,2,ba)=FX(:,col).*POOL_local_BA(:,2,ba);                          % INCOME for the DAY
    POOL_dollars_BA(2:T,1,ba)=           POOL_dollars_BA(1:end-1,4,ba);                      % BEG of DAY BALANCE
    POOL_dollars_BA(:,3,ba)=POOL_dollars_BA(:,4,ba)-sum(POOL_dollars_BA(:,1:2,ba),2); % CAPITAL GAIN
end
clear CONVERT_match_TF col c ba

%% ACCRUAL ACCOUNTS: CONVERT to USD
for aa=1:AA
    for c=1:CC
        CONVERT_match_TF = strcmp(POOL_ccy_AA(aa),FX_ccy(c));
        if CONVERT_match_TF==1; col=c; end
    end
                         POOL_dollars_AA( : ,09,aa)=FX(:,col).*POOL_local_AA(:,5,aa);                                  % BALANCE before REDEMPTIONs EFFECT
                         POOL_dollars_AA( : ,10,aa)=0;                                                                 % REDEMPTIONs EFFECT
                         POOL_dollars_AA( : ,11,aa)=POOL_dollars_AA( : ,9,aa);                                         % END
    switch POOL_CGIswitch_AA{aa}
        case 'INCOME';   POOL_dollars_AA( : ,2,aa)=FX(:,col).*POOL_local_AA(:,2,aa);                                   % ACCRUAL (income)
        case 'CAP.GN';   POOL_dollars_AA( : ,3,aa)=FX(:,col).*POOL_local_AA(:,2,aa);                                   % ACCRUAL (CapGain)
    end
                         POOL_dollars_AA( : ,5,aa)=FX(:,col).*POOL_local_AA(:,3,aa);                                   % REVERSAL (as hits POOL MAIN)
                         POOL_dollars_AA( : ,6,aa)=FX(:,col).*POOL_local_AA(:,3,aa);                                   % REVERSAL (as hits POOL   AA) same as above (for Symmetry with MNGR and CLTS variables)
    switch POOL_CGIswitch_AA{aa}
        case 'INCOME';   POOL_dollars_AA( : ,7,aa)=FX(:,col).*POOL_local_AA(:,4,aa);                                   % WRITE-OFF (income)
        case 'CAP.GN';   POOL_dollars_AA( : ,8,aa)=FX(:,col).*POOL_local_AA(:,4,aa);                                   % WRITE-OFF (CapGain)
    end
                         POOL_dollars_AA(2:T,1,aa)=POOL_dollars_AA(1:T-1,end,aa);                                      % BEG
                         POOL_dollars_AA( : ,4,aa)=POOL_dollars_AA(:,9,aa)-sum(POOL_dollars_AA(:,[1 2 3 5 7 8],aa),2); % FX-TRANSLATION (CapGain)
end
clear CONVERT_match_TF col c a

%% MANAGER
% SUB and RED (in USD)
[MNGR_SubRed,MNGR_SubRed_text]=xlsread('\data\MANAGER_SubRed_NEW.xlsx');

 MNGR_Deposit      = MNGR_SubRed(:,06);
 MNGR_Payouts      = MNGR_SubRed(:,07);
 MNGR_SubRed       = MNGR_SubRed(:,1:4);
 
 MNGR_name         = MNGR_SubRed_text{02,2};
 MNGR_Identifier   = MNGR_SubRed_text{03,2};
 MNGR_TaxResidence = MNGR_SubRed_text{04,2};
 MNGR_Email        = MNGR_SubRed_text{05,2};
 MNGR_Address      = MNGR_SubRed_text{06,2};
 MNGR_TIN          = MNGR_SubRed_text{07,2};
 MNGR_PlaceBirth   = MNGR_SubRed_text{08,2};
 MNGR_DateBirth    = MNGR_SubRed_text{09,2};
 MNGR_Phone        = MNGR_SubRed_text{10,2};
 
clear MNGR_SubRed_text
%% CLIENTS information

[CLTS_SubRed_OLD,CLTS_SubRed_text]=xlsread('\data\CLIENTS_Information.xlsx');

CLTS_SeriesNumber = CLTS_SubRed_text(01,2:4:end); S=size(CLTS_SeriesNumber,2); % NUMBER of CLIENT SERIES
CLTS_Names        = CLTS_SubRed_text(02,2:4:end);
CLTS_Identifier   = CLTS_SubRed_text(03,2:4:end);
CLTS_TaxResidence = CLTS_SubRed_text(04,2:4:end);
CLTS_finderYN     = CLTS_SubRed_text(05,2:4:end);
CLTS_emails       = CLTS_SubRed_text(06,2:4:end);
CLTS_Bloomberg    = CLTS_SubRed_text(07,2:4:end);
CLTS_Address      = CLTS_SubRed_text(08,2:4:end);
CLTS_TIN          = CLTS_SubRed_text(09,2:4:end);
CLTS_PlaceBirth   = CLTS_SubRed_text(10,2:4:end);
CLTS_DateBirth    = CLTS_SubRed_text(11,2:4:end);
CLTS_Phone        = CLTS_SubRed_text(12,2:4:end);

clear CLTS_SubRed_text

%% CLIENTS NEW
%  SUB and RED
%  DEPOSIT and PAYOUTS

[CLTS_SubRed_NEW,CLTS_SubRed_text]=xlsread('\data\CLIENTS_SubRed_NEW_NEW.xlsx');

CLTS_SubRed_NEW  = CLTS_SubRed_NEW(:,1:2:end);
CLTS_SubRed_DESC = CLTS_SubRed_text(3:end,1);
CLTS_SubRed_TEXT = CLTS_SubRed_text(3:end,2:2:end);

R=3; % R is maximal # of REDEMPTIONS in a given series. R may need to increase in the future.
P=3; % P is maximal # of PAYOUTS     in a given series. P may need to increase in the future.
          
CLTS_SubRed         = zeros(T,5,S);         % SUB.RED in USD
CLTS_Deposit        = zeros(T,1,S);         % DEPOSIT in USD
CLTS_Payouts        = zeros(T,1,S);         % PAYOUTS in USD

CLTS_dates_DEP_tinT = zeros(1,S);           % DEPOSIT      DATE (t in [1,T]) Single (1) deposit      per Series (manually aggregate and convert to dollars if multiple deposits).
CLTS_dates_SUB_tinT = zeros(1,S);           % SUBSCRIPTION DATE (t in [1,T]) Single (1) subscription per Series.
CLTS_dates_RED_tinT = zeros(R,S);           % REDEMPTION   DATE (t in [1,T]) Redemptions (R)
CLTS_dates_PAY_tinT = zeros(P,S);           % PAYOUT       DATE (t in [1,T]) Payouts     (P)

CLTS_dates_DEP_dtnm = zeros(1,S);           % DEPOSIT      DATE (in Datenum) Single (1) deposit only (manually aggregate if multiple deposits)
CLTS_dates_SUB_dtnm = zeros(1,S);           % SUBSCRIPTION DATE (in Datenum) Single (1) subscription per Series.
CLTS_dates_RED_dtnm = zeros(R,S);           % REDEMPTION   DATE (in Datenum) Upto R (3) redemptions.
CLTS_dates_PAY_dtnm = zeros(P,S);           % PAYOUT       DATE (in Datenum) Upto P (3) payouts.

% Reshape and Auxilliary variables
for s=1:S
    
    % Deposits
        dt1 = datenum(CLTS_SubRed_TEXT{1,s},'dd/mm/yyyy');          CLTS_dates_DEP_dtnm(s) = dt1;
        dt2 = find(DATES(:,1)==dt1);                                CLTS_dates_DEP_tinT(s) = dt2; 

        CLTS_Deposit(dt2,1,s)=CLTS_SubRed_NEW(1,s); % Deposit is in USD (after conversion)
    
        clear dt1 dt2
    
    % Subscription
        dt1 = datenum(CLTS_SubRed_TEXT{2,s},'dd/mm/yyyy');          CLTS_dates_SUB_dtnm(s) = dt1;
        dt2 = dt1 - DATES(1,1) + 1;                                 CLTS_dates_SUB_tinT(s) = dt2;
                                                                    CLTS_SubRed(dt2,1,s) = CLTS_SubRed_NEW(2,s); % Subscription in USD
                                                                    CLTS_SubRed(dt2,2,s) = CLTS_SubRed_NEW(3,s); % SubCost      in USD
    
        clear dt1 dt2
    
    % Redemptions
    for r=1:R 
        if ~strcmp(CLTS_SubRed_TEXT{4-1+r,s},'NA')
            
            dt1 = datenum(CLTS_SubRed_TEXT{4-1+r,s},'dd/mm/yyyy');  CLTS_dates_RED_dtnm(r,s) = dt1;
            dt2 = find(DATES(:,1)==dt1);                            CLTS_dates_RED_tinT(r,s) = dt2;
                                                                    CLTS_SubRed(dt2,4,s) = CLTS_SubRed_NEW(4-1+r,s); % Redemption   in USD
                                                                    CLTS_SubRed(dt2,3,s) = CLTS_SubRed_NEW(7-1+r,s); % RedCost      in USD
        end
        clear dt1 dt2
    end
    
    % Payouts
    for p=1:P
        if ~strcmp(CLTS_SubRed_TEXT{10-1+p,s},'NA')
            
            dt1 = datenum(CLTS_SubRed_TEXT{10-1+p,s},'dd/mm/yyyy'); CLTS_dates_PAY_dtnm(p,s) = dt1;
            dt2 = find(DATES(:,1)==dt1);                            CLTS_dates_PAY_tinT(p,s) = dt2;
                                                                    CLTS_Payouts(dt2,1,s)    = CLTS_SubRed_NEW(10-1+p,s); % Payout in USD
        end
        clear dt1 dt2
    end
end

% CUMULATIVE SUB/RED for each CLIENT SERIES (for PERF.FEE calculation)
for s=1:S
    for t=1:T
        if  t==1; CLTS_SubRed(t,5,s)=cumsum(CLTS_SubRed(1:t,1,s)); end 
        if  t >1; CUMSUM_sub=cumsum(CLTS_SubRed(1:t  ,1,s));
                  CUMSUM_red=cumsum(CLTS_SubRed(1:t-1,4,s));
                  CLTS_SubRed(t,5,s)=CUMSUM_sub(end)+CUMSUM_red(end); % CUMULATIVE SUB/RED for each CLIENT SERIES (for PERF.FEE calculation)
        end
    end
end

% Redemptions 0/1 Time-Series - Relevant for AA calculations
RedemptionsSeries=zeros(T,S);
for s=1:S
    for t=1:T
        if CLTS_SubRed(t,4,s)<0 || t==CLTS_dates_RED_tinT(R,s) % Partial Redemption (amount known) OR Final Redemption (amount unknown until calculated)
           RedemptionsSeries(t,s)=1;
        end
    end
end

% Series Fully Paid Out
CLTS_FullyPaidOut = find(CLTS_dates_PAY_tinT(P,:)>0);

clear CLTS_SubRed_text CUMSUM_sub CUMSUM_red

%% COSTS SET-UP (20k) (TRANSFER from CLIENTS to MANAGER - NO EFFECT on POOL)
[POOL_SetupCost,POOL_SetupCost_text]=xlsread('\data\POOL_setup_COSTS.xlsx');
 POOL_SetupCost=POOL_SetupCost(:,:);
clear POOL_SetupCost_text

%% Client Fee Parameters
[CLTS_ParaFee,CLTS_ParaFee_text]=xlsread('\data\CLIENTS_ParaFee.xlsx');
 CLTS_ParaFee_MGMT = CLTS_ParaFee(1,1:S) ;
 CLTS_ParaFee_PERF = CLTS_ParaFee(2,1:S) ;
clear CLTS_ParaFee CLTS_ParaFee_text

%% Finders Fee Parameters
[CLTS_ParaFinder,CLTS_ParaFinder_text]=xlsread('\data\CLIENTS_ParaFinder.xlsx'); 

 CLTS_ParaFinder_rates = CLTS_ParaFinder([3 5],1:end);
 CLTS_ParaFinder_dates = CLTS_ParaFinder_text([6 8 10],2:end);
 CLTS_ParaFinder_names = CLTS_ParaFinder_text(4,2:end);
 CLTS_ParaFinder_sries = CLTS_ParaFinder(1,1:end);
 
 CLTS_ParaFinder_rates = reshape(CLTS_ParaFinder_rates,2,3,S);
 CLTS_ParaFinder_dates = reshape(CLTS_ParaFinder_dates,3,3,S);
 CLTS_ParaFinder_names = reshape(CLTS_ParaFinder_names,1,3,S);
 CLTS_ParaFinder_sries = reshape(CLTS_ParaFinder_sries,1,3,S); 

 F=S;
clear CLTS_ParaFinder CLTS_ParaFinder_text

%% Display LOADING XLSX FILES COMPLETED
disp('LOADING XLSX FILES COMPLETED')

%% ZEROS
MNGR_dollars        =zeros(T,27);
MNGR_dollars_AA     =zeros(T,11,AA);
MNGR_dollars_EA     =zeros(T,06,EA);
MNGR_dollars_HC_sub =zeros(T,04);
MNGR_dollars_HC_red =zeros(T,04);
MNGR_percent        =zeros(T,02);
MNGR_percent_AA     =zeros(T,02,AA);

CLTS_dollars        =zeros(T,27,S);
CLTS_dollars_AA     =zeros(T,11,S,AA);
CLTS_dollars_EA     =zeros(T,06,S,EA);
CLTS_dollars_HC_sub =zeros(T,04,S);
CLTS_dollars_HC_red =zeros(T,04,S);
CLTS_percent        =zeros(T,02,S);
CLTS_percent_AA     =zeros(T,02,S,AA);

POOL_dollars        =zeros(T,27);
POOL_dollars_HC_sub =zeros(T,04);
POOL_dollars_HC_red =zeros(T,04);
POOL_SubRed         =zeros(T,04);
POOL_Deposit        =zeros(T,01);
POOL_Payouts        =zeros(T,01);
POOL_percent_BA     =zeros(T,02,BA);
POOL_percent_AA     =zeros(T,02,AA);
POOL_percent_EA     =zeros(T,02,EA);
POOL_percent_HC_sub =zeros(T,02);
POOL_percent_HC_red =zeros(T,02);

CLTS_MgmtFee        =zeros(T,02,S);
CLTS_PerfFee        =zeros(T,05,S);
CLTS_SeriesClosed   =cell(1,S);

MNGR_MgmtFee        =zeros(T,03);
MNGR_PerfFee        =zeros(T,03);

FNDR_dollars_MGMT     =zeros(T,F,S);
FNDR_dollars_PERF     =zeros(T,F,S);
FNDR_dollars_MGMT_sum =zeros(T,1,S);
FNDR_dollars_PERF_sum =zeros(T,1,S);

%% MAIN LOOP

for t=1:T

%% BEGINNING BALANCE
%  t==1 or t==CLTS_dates_SUB_tinT(s)
                                                      % MNGR
                if t==1;                                MNGR_dollars(t,01)          = 0; end                                                                % MNGR MAIN
                if t==1;                                MNGR_dollars_HC_sub(t,01)   = 0; end                                                                % MNGR CUSTODY SUB
                if t==1;                                MNGR_dollars_HC_red(t,01)   = 0; end                                                                % MNGR CUSTODY RED
                if t==1;        for aa=1:AA;            MNGR_dollars_AA(t,01,aa)    = 0; end; end                                                           % MNGR AA
                if t==1;        for ea=1:EA;            MNGR_dollars_EA(t,01,ea)    = 0; end; end                                                           % MNGR EA
                if t==1;        for aa=1:AA;            MNGR_percent_AA(t,01,aa)    = 0; end; end	                                                        % MNGR AA (%) in AA of POOL
                
                                                      % CLTS
for s=1:S;      if t==min(CLTS_dates_DEP_tinT(s),CLTS_dates_SUB_tinT(s));     CLTS_dollars_HC_sub(t,01,s) = 0; end; end                                                           % CLTS HeldCust SUB ($)
for s=1:S;      if t==CLTS_dates_SUB_tinT(s)
                                                        CLTS_dollars(t,01,s)        = 0;                                                                    % CLTS MAIN
                                                        CLTS_dollars_HC_red(t,01,s) = 0;                                                                    % CLTS CUSTODY RED
                                for aa=1:AA;            CLTS_dollars_AA(t,01,s,aa)  = 0; end                                                                % CLTS AA
                                for ea=1:EA;            CLTS_dollars_EA(t,01,s,ea)  = 0; end                                                                % CLTS EA
                                for aa=1:AA;            CLTS_percent_AA(t,01,s,aa)  = 0; end                                                                % CLTS AA (%) in AA of POOL
                end
end
                                                      % POOL
                if t==1;                                POOL_dollars(t,01)          = 0; end                                                                % POOL MAIN
                if t==1;                                POOL_dollars_HC_sub(t,01)   = 0; end                                                                % POOL HeldCust SUB
                if t==1;                                POOL_dollars_HC_red(t,01)   = 0; end                                                                % POOL HeldCust RED
                if t==1;        for ba=1:BA;            POOL_percent_BA(t,01,ba)    = 0; end; end                                                           % POOL BA (%) in POOL MAIN
                if t==1;        for aa=1:AA;            POOL_percent_AA(t,01,aa)    = 0; end; end                                                           % POOL AA (%) in POOL MAIN
                if t==1;        for ea=1:EA;            POOL_percent_EA(t,01,ea)    = 0; end; end                                                           % POOL EA (%) in POOL MAIN
                if t==1;                                POOL_percent_HC_sub(t,01)   = 0; end                                                                % POOL HC sub (%) in POOL MAIN
                if t==1;                                POOL_percent_HC_red(t,01)   = 0; end                                                                % POOL HC red (%) in POOL MAIN

%% BEGINNING BALANCE
%  t>1 or t>CLTS_dates_SUB_tinT(s)
                                                      % MNGR
                if t>1;                                 MNGR_dollars(t,01)          = MNGR_dollars(t-1,end);       end                                      % MNGR MAIN
                if t>1;                                 MNGR_dollars_HC_sub(t,01)   = MNGR_dollars_HC_sub(t-1,end);  end                                    % MNGR CUSTODY SUB
                if t>1;                                 MNGR_dollars_HC_red(t,01)   = MNGR_dollars_HC_red(t-1,end);  end                                    % MNGR CUSTODY RED
                if t>1;         for aa=1:AA;            MNGR_dollars_AA(t,01,aa)    = MNGR_dollars_AA(t-1,end,aa); end; end                                 % MNGR AA
                if t>1;         for ea=1:EA;            MNGR_dollars_EA(t,01,ea)    = MNGR_dollars_EA(t-1,end,ea); end; end                                 % MNGR EA
                if t>1;         for aa=1:AA;            MNGR_percent_AA(t,01,aa)    = MNGR_percent_AA(t-1,end,aa); end; end                                 % MNGR AA (%) in AA of POOL

                                                      % CLTS
for s=1:S;      if t>min(CLTS_dates_DEP_tinT(s),CLTS_dates_SUB_tinT(s));   CLTS_dollars_HC_sub(t,01,s) = CLTS_dollars_HC_sub(t-1,end,s);   end; end         % CLTS CUSTODY SUB
for s=1:S;      if t>CLTS_dates_SUB_tinT(s)
                                                        CLTS_dollars(t,01,s)        = CLTS_dollars(t-1,end,s);                                              % CLTS MAIN
                                                        CLTS_dollars_HC_red(t,01,s) = CLTS_dollars_HC_red(t-1,end,s);                                       % CLTS CUSTODY RED
                                for aa=1:AA;            CLTS_dollars_AA(t,01,s,aa)  = CLTS_dollars_AA(t-1,end,s,aa); end                                    % CLTS AA
                                for ea=1:EA;            CLTS_dollars_EA(t,01,s,ea)  = CLTS_dollars_EA(t-1,end,s,ea); end                                    % CLTS EA
                                for aa=1:AA;            CLTS_percent_AA(t,01,s,aa)  = CLTS_percent_AA(t-1,end,s,aa); end                                    % CLTS AA (%) in AA of POOL
                end
end
                                                      % POOL
                if t>1;                                 POOL_dollars(t,01)          = POOL_dollars(t-1,end);        end                                      % POOL MAIN
                if t>1;                                 POOL_dollars_HC_red(t,01)   = POOL_dollars_HC_red(t-1,end); end                                     % POOL CUSTODY RED
                if t>1;                                 POOL_dollars_HC_sub(t,01)   = POOL_dollars_HC_sub(t-1,end); end                                     % POOL CUSTODY SUB
                if t>1;         for ba=1:BA;            POOL_percent_BA(t,01,ba)    = POOL_percent_BA(t-1,end,ba);  end; end                                 % POOL BA (%) in POOL MAIN
                if t>1;         for aa=1:AA;            POOL_percent_AA(t,01,aa)    = POOL_percent_AA(t-1,end,aa);  end; end                                 % POOL AA (%) in POOL MAIN
                if t>1;         for ea=1:EA;            POOL_percent_EA(t,01,ea)    = POOL_percent_EA(t-1,end,ea);  end; end                                 % POOL EA (%) in POOL MAIN
                if t>1;                                 POOL_percent_HC_sub(t,01)   = POOL_percent_HC_sub(t-1,end); end                                     % POOL HC sub (%) in POOL MAIN
                if t>1;                                 POOL_percent_HC_red(t,01)   = POOL_percent_HC_red(t-1,end); end                                     % POOL HC red (%) in POOL MAIN

%% SUB.RED
%  POOL aggregate
                                                        POOL_SubRed(t,01)           = MNGR_SubRed(t,1) + sum(CLTS_SubRed(t,1,:),3);                         % POOL Subs
                                                        POOL_SubRed(t,02)           = MNGR_SubRed(t,2) + sum(CLTS_SubRed(t,2,:),3);                         % POOL SubCosts
                                                        POOL_SubRed(t,03)           = MNGR_SubRed(t,3) + sum(CLTS_SubRed(t,3,:),3);                         % POOL RedCosts
                                                      % POOL_SubRed(t,04)                                                                                   % POOL Reds kept for later because FINAL REDEMPTION must be calculated at the END of the day.
%% SUBSCRIPTIONS
                                                        MNGR_dollars(t,02)          = MNGR_SubRed(t,1);                                                     % MNGR Subs
for s=1:S;  if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars(t,02,s)        = CLTS_SubRed(t,1,s); end; end                                          % CLTS Subs
                                                        POOL_dollars(t,02)          = POOL_SubRed(t,1);                                                     % POOL Subs

%% SUB COSTS
                                                        MNGR_dollars(t,03)          = MNGR_SubRed(t,2);                                                     % MNGR SubCosts
for s=1:S;  if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars(t,03,s)        = CLTS_SubRed(t,2,s); end; end                                          % CLTS SubCosts
                                                        POOL_dollars(t,03)          = POOL_SubRed(t,2);                                                     % POOL SubCosts
                                                        
%% HELD in CUSTODY
%  SUBSCRIPTIONS
                                                        MNGR_dollars_HC_sub(t,03)   = MNGR_SubRed(t,1);                                                     % SUBS HeldCust MNGR
for s=1:S;	if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars_HC_sub(t,03,s) = CLTS_SubRed(t,1,s); end; end                                          % SUBS HeldCust CLTS
                                                        POOL_dollars_HC_sub(t,03)   = POOL_SubRed(t,1);                                                     % SUBS HeldCust POOL
 
%% BALANCE for OWNERSHIP (%) in POOL
%  for DAY's (regular) ALLOCATIONS
                                                        MNGR_dollars(t,04)          = sum(MNGR_dollars(t,[1 2 3]  ),2);                                     % MNGR BALANCE (for OWNERSHIP (%) in POOL)
for s=1:S;  if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars(t,04,s)        = sum(CLTS_dollars(t,[1 2 3],s),2);  end; end                           % CLTS BALANCE (for OWNERSHIP (%) in POOL)
                                                        POOL_dollars(t,04)          = sum(POOL_dollars(t,[1 2 3]  ),2);                                     % POOL BALANCE (for OWNERSHIP (%) in POOL)

%% OWNERSHIP (%) in POOL
%  for DAY's (regular) ALLOCATIONS
                                                        MNGR_percent(t,01)          = MNGR_dollars(t,4)   / POOL_dollars(t,4);                              % MANAGER (%) OWNERSHIP in POOL
for s=1:S;	if t>=CLTS_dates_SUB_tinT(s);               CLTS_percent(t,01,s)        = CLTS_dollars(t,4,s) / POOL_dollars(t,4); end; end                     % CLIENTS (%) OWNERSHIP in POOL

%% DAY's (regular) ALLOCATIONS
%  BA INCOME
                                                        POOL_dollars(t,05)          = sum(POOL_dollars_BA(t,2,:),3);                                        % INCOME (aggregate BA allocation) POOL + AddBack RedCosts
                                                        MNGR_dollars(t,05)          = MNGR_percent(t,1)   * POOL_dollars(t,05);                             % INCOME (aggregate BA allocation) MNGR
for s=1:S;  if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars(t,05,s)        = CLTS_percent(t,1,s) * POOL_dollars(t,05); end; end                    % INCOME (aggregate BA allocation) CLTS

%% SUB COSTS
%  Add-Back
                                                        POOL_dollars(t,06)          =-POOL_SubRed(t,2);                                                     % SubCosts Add-Back POOL
                                                        MNGR_dollars(t,06)          = MNGR_percent(t,1)   * POOL_dollars(t,06);                             % SubCosts Add-Back MNGR
for s=1:S;  if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars(t,06,s)        = CLTS_percent(t,1,s) * POOL_dollars(t,06); end; end                    % SubCosts Add-Back CLTS

%% DAY's (regular) ALLOCATIONS
%  BA CAP.GN
                                                        POOL_dollars(t,07)          = sum(POOL_dollars_BA(t,3,:),3);                                        % CAP.GN (aggregate BA allocation) POOL
                                                        MNGR_dollars(t,07)          = MNGR_percent(t,1)   * POOL_dollars(t,07);                             % CAP.GN (aggregate BA allocation) MNGR
for s=1:S;  if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars(t,07,s)        = CLTS_percent(t,1,s) * POOL_dollars(t,07); end; end                    % CAP.GN (aggregate BA allocation) CLTS

%% DEPOSIT AddBack
%  MAIN
                                                                        POOL_Deposit(t,01)          = MNGR_Deposit(t,01) + sum(CLTS_Deposit(t,01,:),3);                     % DEPOSIT         POOL
                                                        
                                                                        POOL_dollars(t,08)          =-POOL_Deposit(t,01);                                                   % DEPOSIT AddBack POOL
                                                                        MNGR_dollars(t,08)          = MNGR_percent(t,01)   * POOL_dollars(t,08);                            % DEPOSIT AddBack MNGR
for s=1:S;	if t>=min(CLTS_dates_DEP_tinT(s),CLTS_dates_SUB_tinT(s));	CLTS_dollars(t,08,s)        = CLTS_percent(t,01,s) * POOL_dollars(t,08); end; end                   % DEPOSIT AddBack CLTS

%% HELD in CUSTODY
%  DEPOSITS
                                                                        MNGR_dollars_HC_sub(t,02)   =-MNGR_Deposit(t,01);                                                   % HeldCust MNGR SideAccount DEPOSITS
for s=1:S;	if t>=min(CLTS_dates_DEP_tinT(s),CLTS_dates_SUB_tinT(s));	CLTS_dollars_HC_sub(t,02,s) =-CLTS_Deposit(t,01,s); end; end                                        % HeldCust CLTS SideAccount DEPOSITS
                                                                        POOL_dollars_HC_sub(t,02)   =-POOL_Deposit(t,01);                                                   % HeldCust POOL SideAccount DEPOSITS
                                                        
%% PAYOUTS
%  CLIENTS: Indicate FINAL PAYOUT and calculate FINAL PAYOUT AMOUNT

for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);         if t==CLTS_dates_PAY_tinT(P,s)                                                    % CLTS indicator for FINAL PAYOUT
                                                        
                                                        CLTS_SeriesClosed{5,s}      = DATES(t,1);                                       % CLTS SeriesClosed FinalPayOut DATENUM
                                                        CLTS_SeriesClosed{6,s}      = datestr(DATES(t,1),'dd-mmm-yyyy');                % CLTS SeriesClosed FinalPayOut DATESTR
                                                        CLTS_SeriesClosed{7,s}      = t;                                                % CLTS SeriesClosed FinalPayOut t
                                                        CLTS_SeriesClosed{8,s}      = CLTS_Payouts(t,01,s);                             % CLTS SeriesClosed FinalPayOut AMOUNT (actual)
                                                        
                                                        CLTS_Payouts(t,01,s)        = sum(CLTS_dollars_HC_red(t,[01 02],s),2);          % CLTS              FinalPayOut AMOUNT (calculated)
                                                      end
                end
end

%% PAYOUTS
%  ADD-BACK
                                                        POOL_Payouts(t,01)          = MNGR_Payouts(t,01) + sum(CLTS_Payouts(t,01,:),3);                     % PAYOUTS POOL
                                                        
                                                        POOL_dollars(t,09)          =-POOL_Payouts(t,01);                                                   % PAYOUTS AddBack POOL
                                                        MNGR_dollars(t,09)          = MNGR_percent(t,01)   * POOL_dollars(t,09);                            % PAYOUTS AddBack MNGR
for s=1:S;	if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars(t,09,s)        = CLTS_percent(t,01,s) * POOL_dollars(t,09); end; end                   % PAYOUTS AddBack CLTS

%% HELD in CUSTODY
%  PAYOUTS
                                                        MNGR_dollars_HC_red(t,03)         =-MNGR_Payouts(t,01);                                             % HeldCust MNGR SideAccount PAYOUTS
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_dollars_HC_red(t,03,s)       =-CLTS_Payouts(t,01,s); end; end                                  % HeldCust CLTS SideAccount PAYOUTS
                                                        POOL_dollars_HC_red(t,03)         =-POOL_Payouts(t,01);                                             % HeldCust POOL SideAccount PAYOUTS

%% AA Accrual Accounts
%  MANAGER
        for aa=1:AA;	switch POOL_CGIswitch_AA{aa}
                                    case 'INCOME';      MNGR_dollars_AA(t,02,aa)    = MNGR_percent(t,1) * POOL_dollars_AA(t,02,aa);                         % ACCRUAL (Income) as hits AA & MAIN
                                    case 'CAP.GN';      MNGR_dollars_AA(t,03,aa)    = MNGR_percent(t,1) * POOL_dollars_AA(t,03,aa);                         % ACCRUAL (CapGan) as hits AA & MAIN
                        end
        end
        for aa=1:AA;    if sum(POOL_dollars_AA(t,1:3,aa),2)~=0
                                                        MNGR_percent_AA(t,01,aa)    = sum(MNGR_dollars_AA(t,1:3,aa),2) / sum(POOL_dollars_AA(t,1:3,aa),2);	% Recalculate Manager's ownership (%) in AA due today's accrual (only thing that changes manager's ownership (%) in AA today, all following items allocated using this %-allocation key)
                        end
        end
        for aa=1:AA;                                    MNGR_dollars_AA(t,04,aa)    = MNGR_percent_AA(t,01,aa) * POOL_dollars_AA(t,04,aa); end              % FX-TRANS as hits AA & MAIN
        for aa=1:AA;                                    MNGR_dollars_AA(t,05,aa)    = MNGR_percent(   t,01)    * POOL_dollars_AA(t,05,aa); end              % REVERSAL as hits      MAIN Ledger
        for aa=1:AA;                                    MNGR_dollars_AA(t,06,aa)    = MNGR_percent_AA(t,01,aa) * POOL_dollars_AA(t,06,aa); end              % REVERSAL as hits AA        Ledger
         
        for aa=1:AA;    switch POOL_CGIswitch_AA{aa}
                                    case 'INCOME';      MNGR_dollars_AA(t,07,aa)    = MNGR_percent_AA(t,01,aa) * POOL_dollars_AA(t,07,aa);                  % WRITE-OFF (Income) as hits AA & MAIN
                                    case 'CAP.GN';      MNGR_dollars_AA(t,08,aa)    = MNGR_percent_AA(t,01,aa) * POOL_dollars_AA(t,08,aa);                  % WRITE-OFF (CapGan) as hits AA & MAIN
                        end
        end
        for aa=1:AA;                                    MNGR_dollars_AA(t,09,aa)    = sum(MNGR_dollars_AA(t,[1 2 3 4 6 7 8],aa),2);   end                   % BALANCE ($) before REDEMPTIONS EFFECT of OTHERS
        for aa=1:AA; if POOL_dollars_AA(t,09,aa)~=0;	MNGR_percent_AA(t,02,aa)    =     MNGR_dollars_AA(t,09,aa) / POOL_dollars_AA(t,09,aa); end; end     % END-BAL (%) before REDEMPTIONS EFFECT of SELF (and OTHERS)
             
                                                        MNGR_dollars(t,10)          = sum(MNGR_dollars_AA(t,02,:),3);                                       % ACCRUAL  (Income) as hits MAIN (agg AA)
                                                        MNGR_dollars(t,11)          = sum(MNGR_dollars_AA(t,03,:),3);                                       % ACCRUAL  (CapGan) as hits MAIN (agg AA)
                                                        MNGR_dollars(t,12)          = sum(MNGR_dollars_AA(t,04,:),3);                                       % FX.TRANS (CapGan) as hits MAIN (agg AA)
                                                        MNGR_dollars(t,13)          = sum(MNGR_dollars_AA(t,05,:),3);                                       % REVERSAL          as hits MAIN (agg AA)
                                                        MNGR_dollars(t,14)          = sum(MNGR_dollars_AA(t,07,:),3);                                       % WRITE-OFF(Income) as hits MAIN (agg AA)
                                                        MNGR_dollars(t,15)          = sum(MNGR_dollars_AA(t,08,:),3);                                       % WRITE-OFF(CapGan) as hits MAIN (agg AA)

%% AA Accrual Accounts
%  CLIENTS
for s=1:S
if t>=CLTS_dates_SUB_tinT(s)
        for aa=1:AA;    switch POOL_CGIswitch_AA{aa}
                                    case 'INCOME';      CLTS_dollars_AA(t,02,s,aa)  = CLTS_percent(t,01,s) * POOL_dollars_AA(t,02,aa);                      % ACCRUAL (Income) as hits AA & MAIN
                                    case 'CAP.GN';      CLTS_dollars_AA(t,03,s,aa)  = CLTS_percent(t,01,s) * POOL_dollars_AA(t,03,aa);                      % ACCRUAL (CapGan) as hits AA & MAIN
                        end
        end
        for aa=1:AA
            if sum(POOL_dollars_AA(t,1:3,aa),2)~=0
                                                        CLTS_percent_AA(t,01,s,aa)  = sum(CLTS_dollars_AA(t,1:3,s,aa),2) / sum(POOL_dollars_AA(t,1:3,aa),2); % Recalculate Clients's ownership (%) in AA due today's accrual (only thing that changes client's ownership (%) in AA today, all following items allocated using this %-allocation key)
            end
        end
        for aa=1:AA;                                    CLTS_dollars_AA(t,04,s,aa)  = CLTS_percent_AA(t,01,s,aa)* POOL_dollars_AA(t,04,aa); end             % FX.TRANS as hits AA & MAIN
        for aa=1:AA;                                    CLTS_dollars_AA(t,05,s,aa)  = CLTS_percent(   t,01,s)   * POOL_dollars_AA(t,05,aa); end             % REVERSAL as hits      MAIN Ledger
        for aa=1:AA;                                    CLTS_dollars_AA(t,06,s,aa)  = CLTS_percent_AA(t,01,s,aa)* POOL_dollars_AA(t,06,aa); end             % REVERSAL as hits AA        Ledger

        for aa=1:AA;    switch POOL_CGIswitch_AA{aa}
                                    case 'INCOME';      CLTS_dollars_AA(t,07,s,aa)  = CLTS_percent_AA(t,01,s,aa)* POOL_dollars_AA(t,07,aa);                 % WRITE-OFF(Income) as hits AA & MAIN
                                    case 'CAP.GN';      CLTS_dollars_AA(t,08,s,aa)  = CLTS_percent_AA(t,01,s,aa)* POOL_dollars_AA(t,08,aa);                 % WRITE-OFF(CapGan) as hits AA & MAIN
                        end
        end
        for aa=1:AA;                                    CLTS_dollars_AA(t,09,s,aa)  = sum(CLTS_dollars_AA(t,[1 2 3 4 6 7 8],s,aa),2); end                   % BALANCE ($) before REDEMPTIONS EFFECT of OTHERS
        for aa=1:AA;    if POOL_dollars_AA(t,09,aa)~=0;	CLTS_percent_AA(t,02,s,aa)  =     CLTS_dollars_AA(t,09,s,aa) / POOL_dollars_AA(t,09,aa); end; end	% END-BAL (%) before REDEMPTIONS EFFECT of SELF (and OTHERS)

                                                        CLTS_dollars(t,10,s)        = sum(CLTS_dollars_AA(t,02,s,:),4);                                     % ACCRUAL  (Income) as hits MAIN (agg AA)
                                                        CLTS_dollars(t,11,s)        = sum(CLTS_dollars_AA(t,03,s,:),4);                                     % ACCRUAL  (CapGan) as hits MAIN (agg AA)
                                                        CLTS_dollars(t,12,s)        = sum(CLTS_dollars_AA(t,04,s,:),4);                                     % FX.TRANS (CapGan) as hits MAIN (agg AA)
                                                        CLTS_dollars(t,13,s)        = sum(CLTS_dollars_AA(t,05,s,:),4);                                     % REVERSAL          as hits MAIN (agg AA)
                                                        CLTS_dollars(t,14,s)        = sum(CLTS_dollars_AA(t,07,s,:),4);                                     % WRITE-OFF(Income) as hits MAIN (agg AA)
                                                        CLTS_dollars(t,15,s)        = sum(CLTS_dollars_AA(t,08,s,:),4);                                     % WRITE-OFF(CapGan) as hits MAIN (agg AA)
end
end
%% AA Accrual Accounts
%  POOL
                                                        POOL_dollars(t,10)          = sum(POOL_dollars_AA(t,02,:),3);                                       % ACCRUAL   (Income)    (agg AA)
                                                        POOL_dollars(t,11)          = sum(POOL_dollars_AA(t,03,:),3);                                       % ACCRUAL   (CapGan)    (agg AA)
                                                        POOL_dollars(t,12)          = sum(POOL_dollars_AA(t,04,:),3);	                                    % FX.TRANS              (agg AA)
                                                        POOL_dollars(t,13)          = sum(POOL_dollars_AA(t,05,:),3);	                                    % REVERSAL as hits MAIN (agg AA)
                                                        POOL_dollars(t,14)          = sum(POOL_dollars_AA(t,07,:),3);                                       % WRITE-OFF (Income)    (agg AA)
                                                        POOL_dollars(t,15)          = sum(POOL_dollars_AA(t,08,:),3);	                                    % WRITE-OFF (CapGan)    (agg AA)
                      
%% EA Expense Accounts
%  MANAGER
                                for ea=1:EA;            MNGR_dollars_EA(t,02,ea)    =                  0 * POOL_dollars_EA(t,02,ea); end                     % DEPRECIATION (EA)
                                for ea=1:EA;            MNGR_dollars_EA(t,03,ea)    =  MNGR_percent(t,1) * POOL_dollars_EA(t,03,ea); end                     % CAPITALIZATN as hits MAIN
                                for ea=1:EA;            MNGR_dollars_EA(t,04,ea)    =                  0 * POOL_dollars_EA(t,04,ea); end                     % CAPITALIZATN as hits EA
                      
                                                        MNGR_dollars(t,16)          = sum(MNGR_dollars_EA(t,02,:),3);                                       % DEPRECIATION              (agg EA)
                                                        MNGR_dollars(t,17)          = sum(MNGR_dollars_EA(t,03,:),3);                                       % CAPITALIZATN as hits MAIN (agg EA)
%% EA Expense Accounts
%  CLIENTS
for s=1:S
if t>=CLTS_dates_SUB_tinT(s)
                                for ea=1:EA;            CLTS_dollars_EA(t,02,s,ea)  = CLTS_percent(t,1,s)/(1-MNGR_percent(t,1)) * POOL_dollars_EA(t,02,ea); end     % DEPRECIATION (EA)
                                for ea=1:EA;            CLTS_dollars_EA(t,03,s,ea)  = CLTS_percent(t,1,s)                       * POOL_dollars_EA(t,03,ea); end     % CAPITALIZATN as hits MAIN
                                for ea=1:EA;            CLTS_dollars_EA(t,04,s,ea)  = CLTS_percent(t,1,s)/(1-MNGR_percent(t,1)) * POOL_dollars_EA(t,04,ea); end     % CAPITALIZATN as hits EA

                                                        CLTS_dollars(t,16,s)        = sum(CLTS_dollars_EA(t,02,s,:),4);                                             % DPRCDEPRECIATIONTON       (agg EA)
                                                        CLTS_dollars(t,17,s)        = sum(CLTS_dollars_EA(t,03,s,:),4);                                             % CAPITALIZATN as hits MAIN (agg EA)
end
end
%% EA Expense Accounts
%  POOL
                                                        POOL_dollars(t,16)          = sum(POOL_dollars_EA(t,02,:),3);                                       % DEPRECIATION (agg EA)
                                                        POOL_dollars(t,17)          = sum(POOL_dollars_EA(t,03,:),3);                                       % CAPITALIZATN (agg EA)
                                                        
%% RED COSTS
                                                        MNGR_dollars(t,18)          = MNGR_SubRed(t,03);                                                    % RedCosts MNGR
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_dollars(t,18,s)        = CLTS_SubRed(t,03,s); end; end                                         % RedCosts CLTS
                                                        POOL_dollars(t,18)          = POOL_SubRed(t,03);                                                    % RedCosts POOL

%% RED COSTS
%  Add-Back
                                                        POOL_dollars(t,19)          =-POOL_SubRed(t,03);                                                    % RedCosts Add-Back POOL
                                                        MNGR_dollars(t,19)          = MNGR_percent(t,1)   * POOL_dollars(t,19);                             % RedCosts Add-Back MNGR
for s=1:S;  if t>=CLTS_dates_SUB_tinT(s);               CLTS_dollars(t,19,s)        = CLTS_percent(t,1,s) * POOL_dollars(t,19); end; end                    % RedCosts Add-Back CLTS

%% BALANCE
% before SetupCosts, MgmtFees, PerfFees, Redemptions.
                                                        POOL_dollars(t,20)          = sum(POOL_dollars(t,4:19  ),2);                                        % POOL BALANCE,  before                                 Redemptions, for calculating End-of-Day PRICES of Shadow Shares of the POOL,
                                                        MNGR_dollars(t,20)          = sum(MNGR_dollars(t,4:19  ),2);                                        % MNGR BALANCE,  before SetupCosts, MgmtFees, PerfFees, Redemptions, for calculating End-of-Day PRICES of        Shares at which to issue COMPENSATION SHARES and at which to redeem shares.
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_dollars(t,20,s)        = sum(CLTS_dollars(t,4:19,s),2); end; end                               % CLTS BALANCE,  before SetupCosts, MgmtFees, PerfFees, Redemptions, for symmetry with POOL and MNGR variables.

%% SetupCosts
                                                        MNGR_dollars(t,21)          =-POOL_SetupCost(t);                                                        % MNGR SetupCosts
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_dollars(t,21,s)        = POOL_SetupCost(t) * CLTS_percent(t,1,s)/(1-MNGR_percent(t,1)); end; end   % CLTS SetupCosts
                                                        POOL_dollars(t,21)          = 0;                                                                        % POOL SetupCosts, for symmetry with MNGR and CLTS variables.
%% Management Fees
%  CLIENTS
for s=1:S
if t>=CLTS_dates_SUB_tinT(s)    
                                                        CLTS_MgmtFee(t,01,s) = sum(CLTS_dollars(t,[20 21],s),2);                                            % BALANCE for MANAGMENT FEE calculation (includes the deduction of SetupCosts)
                                                        CLTS_MgmtFee(t,02,s) = 1/12 * DATES(t,2) * CLTS_ParaFee_MGMT(s) * CLTS_MgmtFee(t,01,s);             % MANAGMENT FEE in MgmtFee variable
                                                        CLTS_dollars(t,22,s) = -CLTS_MgmtFee(t,02,s);                                                       % MANAGMENT FEE in MAIN    variable
end
end
%% Performance Fees
%  CLIENTS
for s=1:S
if t>=CLTS_dates_SUB_tinT(s)
                                                        CLTS_PerfFee(t,01,s) = CLTS_MgmtFee(t,01,s) - CLTS_MgmtFee(t,02,s);                                 % BALANCE for PERFORMANCE FEE calculation (includes the deduction of MgmtFee)
                                                        CLTS_PerfFee(t,02,s) = CLTS_PerfFee(t,01,s) - CLTS_SubRed(t,05,s);                                  % Cumulative Gain, adjusted for subscriptions and redemptions
if t==CLTS_dates_SUB_tinT(s);                                      CLTS_PerfFee(t,03,s) = 0; end                                                                       % Adj Cumulative GAIN s.t. WATERMARK
if t==CLTS_dates_SUB_tinT(s);                                      CLTS_PerfFee(t,04,s) = 0; end                                                                       % Adj Cumulative GAIN s.t. WATERMARK
                      if DATES(t,2)==1;              if CLTS_PerfFee(t,2,s) > CLTS_PerfFee(t-1,3,s)                                                         % Adj Cumulative GAIN s.t. WATERMARK
                                                        CLTS_PerfFee(t,3,s) = CLTS_PerfFee(t  ,2,s);                                                        % Adj Cumulative GAIN s.t. WATERMARK
                                                   else CLTS_PerfFee(t,3,s) = CLTS_PerfFee(t-1,3,s);    end                                                 % Adj Cumulative GAIN s.t. WATERMARK
                                                   else CLTS_PerfFee(t,3,s) = CLTS_PerfFee(t-1,3,s);    end                                                 % Adj Cumulative GAIN s.t. WATERMARK
                                                        CLTS_PerfFee(t,4,s) = max(CLTS_PerfFee(t,3,s) - CLTS_PerfFee(t-1,3,s),0);                           % COMPENSABLE GAIN
                                                        CLTS_PerfFee(t,5,s) = DATES(t,2) * CLTS_ParaFee_PERF(s) * CLTS_PerfFee(t,4,s);                      % PERFORMANCE FEE
                                                        CLTS_dollars(t,23,s) = -CLTS_PerfFee(t,5,s);                                                        % PERFORMANCE FEE
end
end

%% Finders Fees
%  Finder-CLIENTS
for s=1:S
if t>=CLTS_dates_SUB_tinT(s)
    
                                        for f=1:3
                                            if  any(CLTS_ParaFinder_sries(1,f,s))
                                                
                                                dtnm=datenum(CLTS_ParaFinder_dates(:,f,s),'DD/mm/YYYY');
                                                
                                                if DATES(t,1) >= dtnm(1) && DATES(t,1) <= dtnm(2)
                                                    
                                                    FNDR_dollars_MGMT(t,s,CLTS_ParaFinder_sries(1,f,s)) = CLTS_MgmtFee(t,end,s) * CLTS_ParaFinder_rates(1,f,s); % FinderFee paid by each client (upto three Finders per Client)
                                                    FNDR_dollars_PERF(t,s,CLTS_ParaFinder_sries(1,f,s)) = CLTS_PerfFee(t,end,s) * CLTS_ParaFinder_rates(1,f,s); % FinderFee paid by each client (upto three Finders per Client)
                                                    
                                                end
                                                if DATES(t,1) >  dtnm(2) && DATES(t,1) <= dtnm(3)
                                                    
                                                    FNDR_dollars_MGMT(t,s,CLTS_ParaFinder_sries(1,f,s)) = CLTS_MgmtFee(t,end,s) * CLTS_ParaFinder_rates(2,f,s); % FinderFee paid by each client (upto three Finders per Client)
                                                    FNDR_dollars_PERF(t,s,CLTS_ParaFinder_sries(1,f,s)) = CLTS_PerfFee(t,end,s) * CLTS_ParaFinder_rates(2,f,s); % FinderFee paid by each client (upto three Finders per Client)
                                                    
                                                end
                                            end
                                        end
                                        for f=1:S
                                            FNDR_dollars_MGMT_sum(t,1,f)=sum(FNDR_dollars_MGMT(t,:,f),2);
                                            FNDR_dollars_PERF_sum(t,1,f)=sum(FNDR_dollars_PERF(t,:,f),2);
                                            CLTS_dollars(t,24,f)=FNDR_dollars_MGMT_sum(t,1,f)+FNDR_dollars_PERF_sum(t,1,f); % FINDER received   
                                        end
                                        
end
clear dtnm
end

%% MgmtFee, PerfFee, FndrFee due MANAGER
%  MANAGER
                                                        MNGR_MgmtFee(t,01)= sum(CLTS_MgmtFee(t,end,:),3);           % MGMT fee (gross     due MNGR)
                                                        MNGR_MgmtFee(t,02)=-sum(FNDR_dollars_MGMT_sum(t,1,:),3);    % MGMT fee (aggregate due FNDR)
                                                        MNGR_MgmtFee(t,03)= sum(MNGR_MgmtFee(t,1:2),2);             % MGMT fee (net       due MNGR)

                                                        MNGR_PerfFee(t,01)= sum(CLTS_PerfFee(t,end,:),3);           % PERF fee (gross     due MNGR)
                                                        MNGR_PerfFee(t,02)=-sum(FNDR_dollars_PERF_sum(t,1,:),3);    % PERF fee (aggregate due FNDR)
                                                        MNGR_PerfFee(t,03)= sum(MNGR_PerfFee(t,1:2),2);             % PERF fee (net       due MNGR)

                                                        MNGR_dollars(t,22)          = MNGR_MgmtFee(t,03);                  % MANAGEMENT FEE due manager (NET of FndrFee)
                                                        MNGR_dollars(t,23)          = MNGR_PerfFee(t,03);                  % PERFORMNCE FEE due manager (NET of FndrFee)
                                                        MNGR_dollars(t,24)          = 0;                                   % FINDERS    FEE Left Blank  (MgmtFee and PerfFee are net of FndrFee), for symmetry with CLTS variable.
% POOL
                                                        POOL_dollars(t,22)          = 0;                                   % MANAGEMENT FEE Left Blank for POOL variable, for symmetry with CLTS and MNGR variables.
                                                        POOL_dollars(t,23)          = 0;                                   % PERFORMNCE FEE Left Blank for POOL variable, for symmetry with CLTS and MNGR variables.
                                                        POOL_dollars(t,24)          = 0;                                   % FINDERS    FEE Left Blank for POOL variable, for symmetry with CLTS and MNGR variables.
                                                        
%% BALANCE
%  before REDEMPTION
                                                        MNGR_dollars(t,25)          = sum(MNGR_dollars(t,20:24)  ,2);                                       % MNGR BALANCE, before Redemptions (including Reds add-back), for symmetry with CLTS variable.
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_dollars(t,25,s)        = sum(CLTS_dollars(t,20:24,s),2); end; end                              % CLTS BALANCE, before Redemptions (including Reds add-back), for End-of-Day Share PRICES at which to redeem shares.
                                                        POOL_dollars(t,25)          = sum(POOL_dollars(t,20:24)  ,2);                                       % POOL BALANCE, before Redemptions (including Reds add-back), for symmetry with CLTS variable.
                                                        
%% REDEMPTIONS
%  CLIENTS: Indicate FINAL REDEMPTION and calculate FINAL REDEMPTION AMOUNT
                                                        
                                                        
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);        if CLTS_dates_RED_tinT(R,s)==t                                                                         % CLTS indicator for FINAL REDEMPTIONS, including calculating the FINAL REDEMPTION AMOUNT.
                                                        CLTS_SubRed(t,04,s)         = -CLTS_dollars(t,25,s);                                                % CLTS Calculated FINAL REDEMPTION AMOUNT
                                                        
                                                        CLTS_SeriesClosed{1,s}      = DATES(t,1);                                                           % CLTS SeriesClosed DATENUM
                                                        CLTS_SeriesClosed{2,s}      = datestr(DATES(t,1),'dd-mmm-yyyy');                                    % CLTS SeriesClosed DATESTR
                                                        CLTS_SeriesClosed{3,s}      = t;                                                                    % CLTS SeriesClosed t
                                                        CLTS_SeriesClosed{4,s}      = CLTS_SubRed(t,04,s);                                                  % CLTS SeriesClosed FINAL REDEMPTION AMOUNT
                                                     end
                end
end

%% REDEMPTIONS
%  POOL: Update with FINAL REDEMPTION AMOUNT of CLIENTS
                                                        POOL_SubRed(t,04)           = MNGR_SubRed(t,04) + sum(CLTS_SubRed(t,04,:),3);   % POOL SubRed REDEMPTIONS

%% REDEMPTIONS
%  MNGR, CLTS, POOL
                                                        MNGR_dollars(t,26)          = MNGR_SubRed(t,04);                                % MNGR Redemption
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_dollars(t,26,s)        = CLTS_SubRed(t,04,s); end; end                     % CLTS Redemption
                                                        POOL_dollars(t,26)          = POOL_SubRed(t,04);                                % POOL Redemption
                                                        
%% HELD in CUSTODY
%  REDEMPTIONS
                                                        MNGR_dollars_HC_red(t,02)         = MNGR_SubRed(t,04);                                % HeldCust MNGR SideAccount REDEMPTIONS
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_dollars_HC_red(t,02,s)       = CLTS_SubRed(t,04,s); end; end                     % HeldCust CLTS SideAccount REDEMPTIONS
                                                        POOL_dollars_HC_red(t,02)         = POOL_SubRed(t,04);                                % HeldCust POOL SideAccount REDEMPTIONS

%% ENDING BALANCE
%  MAIN
                                                        MNGR_dollars(t,27)          = sum(MNGR_dollars(t,25:26)  ,2);           % MNGR ENDING BALANCE
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_dollars(t,27,s)        = sum(CLTS_dollars(t,25:26,s),2); end; end  % CLTS ENDING BALANCE
                                                        POOL_dollars(t,27)          = sum(POOL_dollars(t,25:26)  ,2);           % POOL ENDING BALANCE
%% ENDING BALANCE
%  HELD-in-CUSTODY
                                                                     MNGR_dollars_HC_sub(t,04)         = sum(MNGR_dollars_HC_sub(t,01:03)  ,2);           % MNGR HeldCust SUB ($) END
for s=1:S; if t>=min(CLTS_dates_DEP_tinT(s),CLTS_dates_SUB_tinT(s)); CLTS_dollars_HC_sub(t,04,s)       = sum(CLTS_dollars_HC_sub(t,01:03,s),2); end; end  % CLTS HeldCust SUB ($) END
                                                                     POOL_dollars_HC_sub(t,04)         = sum(POOL_dollars_HC_sub(t,01:03)  ,2);           % POOL HeldCust SUB ($) END

                                                                     MNGR_dollars_HC_red(t,04)         = sum(MNGR_dollars_HC_red(t,01:03)  ,2);           % MNGR HeldCust RED END
for s=1:S; if t>=CLTS_dates_SUB_tinT(s);                             CLTS_dollars_HC_red(t,04,s)       = sum(CLTS_dollars_HC_red(t,01:03,s),2); end; end  % CLTS HeldCust RED END
                                                                     POOL_dollars_HC_red(t,04)         = sum(POOL_dollars_HC_red(t,01:03)  ,2);           % POOL HeldCust RED END
%% ENDING BALANCE
%  PERCENT (%) OWNERSHIP in POOL
                                                        MNGR_percent(t,02)          = MNGR_dollars(t,end)   / POOL_dollars(t,end);          % MNGR ENDING OWNERSHIP (%) in POOL
for s=1:S;      if t>=CLTS_dates_SUB_tinT(s);           CLTS_percent(t,02,s)        = CLTS_dollars(t,end,s) / POOL_dollars(t,end); end; end % CLTS ENDING OWNERSHIP (%) in POOL

%% ENDING BALANCE
%  BA,AA,EA (%) in POOL        
                                for ba=1:BA;            POOL_percent_BA(t,2,ba)     = POOL_dollars_BA(t,end,ba) / POOL_dollars(t,end); end  % BA (%) in POOL
                                for aa=1:AA;            POOL_percent_AA(t,2,aa)     = POOL_dollars_AA(t,end,aa) / POOL_dollars(t,end); end  % AA (%) in POOL
                                for ea=1:EA;            POOL_percent_EA(t,2,ea)     = POOL_dollars_EA(t,end,ea) / POOL_dollars(t,end); end  % EA (%) in POOL
                                                        POOL_percent_HC_sub(t,02)   = POOL_dollars_HC_sub(t,end)/ POOL_dollars(t,end);      % HC sub (%) in POOL
                                                        POOL_percent_HC_red(t,02)   = POOL_dollars_HC_red(t,end)/ POOL_dollars(t,end);      % HC sub (%) in POOL

%% EFFECT of REDEMPTIONS
%  on EXPENSE ACCOUNTS (EA)
%  MANAGER & CLIENTS (no effect on POOL)
                                           for ea=1:EA;            MNGR_dollars_EA(t,06,ea)    = 0                                           * POOL_dollars_EA(t,end,ea); end          	% MNGR EA ENDING BALANCE
for s=1:S; if t>=CLTS_dates_SUB_tinT(s);   for ea=1:EA;            CLTS_dollars_EA(t,06,s,ea)  = CLTS_percent(t,2,s) / (1-MNGR_percent(t,2)) * POOL_dollars_EA(t,end,ea); end; end; end	% CLTS EA ENDING BALANCE

                                           for ea=1:EA;            MNGR_dollars_EA(t,05,ea)    = MNGR_dollars_EA(t,end,ea)   - sum(MNGR_dollars_EA(t,[1 2 4],ea),2); end                 	% MNGR EFFECT of SUB/RED
for s=1:S; if t>=CLTS_dates_SUB_tinT(s);   for ea=1:EA;            CLTS_dollars_EA(t,05,s,ea)  = CLTS_dollars_EA(t,end,s,ea) - sum(CLTS_dollars_EA(t,[1 2 4],s,ea),2); end; end; end       % CLTS EFFECT of SUB/RED

%% EFFECT of REDEMPTIONS on AA
%% on SELF in DOLLARS
    
            for aa=1:AA
                if strcmp(POOL_MEMswitch_AA{aa},'REDISTRIBUTES')    
                
                % CLTS ($)
                for s=1:S
                    if t>=CLTS_dates_SUB_tinT(s)
                        if CLTS_dollars(t,25,s)~=0 % If NOT a redeemed series. Not allowed to divide by zero - only relevant if there is a redemption (partial or final).
                           CLTS_dollars_AA(t,10,s,aa) = CLTS_dollars_AA(t,09,s,aa) * CLTS_dollars(t,26,s)/CLTS_dollars(t,25,s);   % REDEMPTION EFFECT
                        end
                    end
                end
                
                % MNGR ($)
                           MNGR_dollars_AA(t,10,aa)=0;                            % REDEMPTION EFFECT of SELF
                         % This is zero because the IM Series is special in
                         % the sense that additional subscriptions are
                         % added to the same Series (as opposed to reults in
                         % the openning of a new Series, and so we cannot
                         % have that Redemptions lead to a reduction in the
                         % AA balances, because we cannot have Subscriptions
                         % symmetrically lead to an increase in the AA balances,
                         % and so this leaves an openning for inconsistency
                         % in that either Redemptions decrease (ok) and
                         % Subscriptions increase (problem), OR, Redemption
                         % decrease(not ok) and Subscriptions don't
                         % increase. So the solution is that neither
                         % Redemptions nor Subscriptions affect the AA
                         % balances, which is fine because the IM Series
                         % will always remain subscribed in a meaningful
                         % way in the Fund.
                         
                % POOL ($, no effect)
                end
            end

%% EFFECT of REDEMPTIONS on AA
%% on SELF in PERCENT
    
    if sum(RedemptionsSeries(t,:))>=1               % If ANY redemption (partial or full).
       RedSer=find(RedemptionsSeries(t,:)==1);      % Redeemed series (can be more than one).
            for aa=1:AA
                if strcmp(POOL_MEMswitch_AA{aa},'REDISTRIBUTES')
                    
                for s=1:S
                    if t>=CLTS_dates_SUB_tinT(s)
                       if POOL_dollars_AA(t,11,aa)~=0
                          CLTS_percent_AA(t,02,s,aa)=sum(CLTS_dollars_AA(t,[9 10],s,aa),2)/(POOL_dollars_AA(t,11,aa)+sum(CLTS_dollars_AA(t,10,RedSer,aa),3));
                       end
                    end
                end
                       if POOL_dollars_AA(t,11,aa)~=0
                          MNGR_percent_AA(t,02,aa)  =sum(MNGR_dollars_AA(t,[9 10],aa),2)  /(POOL_dollars_AA(t,11,aa)+sum(CLTS_dollars_AA(t,10,RedSer,aa),3));
                       end
                end
            end
    end
    
%% EFFECT of REDEMPTIONS on AA
%% on OTHERS in DOLLARS
    
    if sum(RedemptionsSeries(t,:))>=1
       
        RedSer=find(RedemptionsSeries(t,:)==1);  % Find s (in S) of redeeming series (Partial or full, can be more than 1).
        RS=length(RedSer);                       % RS = # of redeeming series.
        
        for aa=1:AA
            if strcmp(POOL_MEMswitch_AA{aa},'REDISTRIBUTES')
            for rs=1:RS
                S_ex_RS{rs}=1:S;
                S_ex_RS{rs}(RedSer(rs))=[];      % S_ex_RS are non-redeeming series.
                RedSer_dollar_REDUCTION_AA_due_self(t,RedSer(rs),aa) = CLTS_dollars_AA(t,10,RedSer(rs),aa);
                                SUM_percent_AA_ex_RS(t,RedSer(rs),aa) = sum(CLTS_percent_AA(t,02,S_ex_RS{rs},aa),3)+MNGR_percent_AA(t,02,aa);
            end
            for rs=1:RS
                if abs(SUM_percent_AA_ex_RS(t,RedSer(rs),aa))>10^(-8)
                    
                    MNGR_dollars_AA(t,10,aa)   = MNGR_dollars_AA(t,10,aa)   - RedSer_dollar_REDUCTION_AA_due_self(t,RedSer(rs),aa) *  MNGR_percent_AA(t,02,aa)   / SUM_percent_AA_ex_RS(t,RedSer(rs),aa);
                    for s=S_ex_RS{rs}
                    CLTS_dollars_AA(t,10,s,aa) = CLTS_dollars_AA(t,10,s,aa) - RedSer_dollar_REDUCTION_AA_due_self(t,RedSer(rs),aa) *  CLTS_percent_AA(t,02,s,aa) / SUM_percent_AA_ex_RS(t,RedSer(rs),aa);
                    end
                end
            end
            end
        end
    end

    %% AA balance END
    for aa=1:AA
                                            MNGR_dollars_AA(t,11,aa)   =sum(MNGR_dollars_AA(t,[9 10],aa),2);
            if POOL_dollars_AA(t,11,aa)~=0; MNGR_percent_AA(t,02,aa)   =    MNGR_dollars_AA(t,11,aa)   / POOL_dollars_AA(t,11,aa); end
        for s=1:S
                                            CLTS_dollars_AA(t,11,s,aa) =sum(CLTS_dollars_AA(t,[9 10],s,aa),2);
            if POOL_dollars_AA(t,11,aa)~=0; CLTS_percent_AA(t,02,s,aa) =    CLTS_dollars_AA(t,11,s,aa) / POOL_dollars_AA(t,11,aa); end
        end
    end
    
    
    
end % for t=1:T

%% ZEROS for SHARES & PRICES & RETURN
            
POOL_shares = zeros(T,05);
MNGR_shares = zeros(T,08);
CLTS_shares = zeros(T,05,S);

POOL_prices = zeros(T,02);
MNGR_prices = zeros(T,02);
CLTS_prices = zeros(T,02,S);

POOL_return = zeros(T,01);

%% POOL: shadow SHARES and PRICES
for t=1:T
    if t==1;        POOL_shares(t,1) = 0;                    end
    if t==1;        POOL_prices(t,1) = 1000;                 end
    if t >1;        POOL_shares(t,1) = POOL_shares(t-1,end); end
    if t >1;        POOL_prices(t,1) = POOL_prices(t-1,end); end              % BEG
    
                    POOL_shares(t,2) = POOL_dollars(t,2) / POOL_prices(t,1);    % SUBSCRIPTION (Beg-of-Day)
                    POOL_shares(t,3) = sum(POOL_shares(t,[1 2]),2);           % SHARES BALANCE (before Redemptions)
                    
                    POOL_prices(t,2) = POOL_dollars(t,20) / POOL_shares(t,3); % PRICE END-of-DAY
                    POOL_return(t,1) = POOL_prices(t,2) / POOL_prices(t,1)-1; % Daily RETURN
                    
                    POOL_shares(t,4) = POOL_dollars(t,26) / POOL_prices(t,2); % REDEMPTION (End-of-Day)
                    POOL_shares(t,5) = sum(POOL_shares(t,3:4),2);             % END
end

% POOL: EXPORT shadow PRICES
filename=strcat('SharePrices_01Jan2015_',datestr(DATES(END,1),'DDmmmYYYY'),'_VGG7S67R1245');
file    =[str2num(datestr(DATES(:,1),'YYYYmmDD')),POOL_prices];
xlswrite(strcat('\reports\',filename),file);


%% MANAGER: SHARES and PRICES
for t=1:T
    if t==1;        MNGR_shares(t,1) = 0;                       end
    if t==1;        MNGR_prices(t,1) = 1000;                    end
    if t >1;        MNGR_shares(t,1) = MNGR_shares(t-1,end); end
    if t >1;        MNGR_prices(t,1) = MNGR_prices(t-1,end); end                 % BEG
    
                    MNGR_shares(t,2) = MNGR_dollars(t,2)/MNGR_prices(t,1);    % SUBSCRIPTION (Beg-of-Day)
                    MNGR_shares(t,3) = sum(MNGR_shares(t,[1 2]),2);              % SHARES BALANCE (before issuance of COMPENSATION SHARES and Redemption of Shares)
                    
                    MNGR_prices(t,2) = MNGR_dollars(t,20) / MNGR_shares(t,3); % PRICE END-of-DAY
%                     MNGR_return(t,1) = MNGR_prices(t,2) / MNGR_prices(t,1)-1; % Daily RETURN
                    
                    MNGR_shares(t,4) = MNGR_dollars(t,21) / MNGR_prices(t,2); % COSTS: set-up
                    MNGR_shares(t,5) = MNGR_dollars(t,22) / MNGR_prices(t,2); % MANAGEMENT
                    MNGR_shares(t,6) = MNGR_dollars(t,23) / MNGR_prices(t,2); % PERFORMANCE
                    MNGR_shares(t,7) = MNGR_dollars(t,26) / MNGR_prices(t,2); % REDEMPTION (End-of-Day)
                    MNGR_shares(t,8) = sum(MNGR_shares(t,3:7),2);                % END
end

%% CLIENTS: SHARES and PRICES
for t=1:T
    for s=1:S
        if t>=CLTS_dates_SUB_tinT(s)
            if t==CLTS_dates_SUB_tinT(s);    CLTS_shares(t,1,s) = 0;                         end
            if t==CLTS_dates_SUB_tinT(s);    CLTS_prices(t,1,s) = 1000;                      end
            if t >CLTS_dates_SUB_tinT(s);    CLTS_shares(t,1,s) = CLTS_shares(t-1,end,s); end
            if t >CLTS_dates_SUB_tinT(s);    CLTS_prices(t,1,s) = CLTS_prices(t-1,end,s); end
        
                                            CLTS_shares(t,2,s) = CLTS_dollars(t,2,s) /CLTS_prices(t,1,s);         % SUBSCRIPTION (Beg-of-Day)
                                            CLTS_shares(t,3,s) = sum(CLTS_shares(t,[1 2],s),2);                      % SHARES BALANCE (before redemptions)
            
                                         if CLTS_shares(t,3,s)~=0
                                            CLTS_prices(t,2,s) = CLTS_dollars(t,25,s) / CLTS_shares(t,3,s);       % END (price)
                                         end
                                         if CLTS_shares(t,3,s)==0
                                            CLTS_prices(t,2,s) = CLTS_prices(t,1,s);                                 % END (price)
                                         end
            
                                            CLTS_shares(t,4,s) = CLTS_dollars(t,26,s)/CLTS_prices(t,2,s);         % REDEMPTION (End-of-Day)
                                         if abs( CLTS_shares(t,3,s) + CLTS_shares(t,4,s) ) < 0.0000001 && CLTS_shares(t,4,s) < 0
                                            CLTS_shares(t,4,s)=-CLTS_shares(t,3,s);
                                         end
                                            CLTS_shares(t,5,s) = sum(CLTS_shares(t,3:4,s),2);                        % END (balance)
        end
    end
end
%% Display Main Loop Completed
disp('MAIN LOOP COMPLETED')
%% CHECKS dollars POOL
CHECKS_dollars       = POOL_dollars -MNGR_dollars -sum(CLTS_dollars,3); % POOL vs MNGR+CLTS 

CHECKS_dollars(:,28) = POOL_dollars(:,001)   - sum(POOL_dollars_BA(:,001,:),3)...
                                             - sum(POOL_dollars_AA(:,001,:),3)...
                                             - sum(POOL_dollars_EA(:,001,:),3)...
                                             - POOL_dollars_HC_sub(:,001)...
                                             - POOL_dollars_HC_red(:,001);    % POOL vs BA+AA+EA +HeldCust   BEG
                                         
CHECKS_dollars(:,29) = POOL_dollars(:,end)   - sum(POOL_dollars_BA(:,end,:),3)...
                                             - sum(POOL_dollars_AA(:,end,:),3)...
                                             - sum(POOL_dollars_EA(:,end,:),3)...
                                             - POOL_dollars_HC_sub(:,end)...
                                             - POOL_dollars_HC_red(:,end);    % POOL vs BA+AA+EA +HeldCust   END
                                         
                                         
[CHECKS_dollars_MAX(1,:),CHECKS_dollars_MAX(2,:)]=max(CHECKS_dollars);
[CHECKS_dollars_MIN(1,:),CHECKS_dollars_MIN(2,:)]=min(CHECKS_dollars);

CHECKS_dollars_MAX_MAX=max(CHECKS_dollars_MAX(1,:));
CHECKS_dollars_MIN_MIN=min(CHECKS_dollars_MAX(1,:));

%% CHECKS percent POOL

CHECKS_percent(:,1) = 1 - MNGR_percent(:,1)-sum(CLTS_percent(:,1,:),3); % BEG (%)
CHECKS_percent(:,2) = 1 - MNGR_percent(:,2)-sum(CLTS_percent(:,2,:),3); % END (%)
CHECKS_percent(:,3) = 1 - sum(POOL_percent_BA(:,2,:),3) ...
                        - sum(POOL_percent_AA(:,2,:),3) ...
                        - sum(POOL_percent_EA(:,2,:),3) ...
                        - POOL_dollars_HC_sub(:,end) ./ POOL_dollars(:,end) ...
                        - POOL_dollars_HC_red(:,end) ./ POOL_dollars(:,end);
                  
[CHECKS_percent_MAX(1,:),CHECKS_percent_MAX(2,:)]=max(CHECKS_percent);
[CHECKS_percent_MIN(1,:),CHECKS_percent_MIN(2,:)]=min(CHECKS_percent);

CHECKS_percent_MAX_MAX=max(CHECKS_percent_MAX(1,:));
CHECKS_percent_MIN_MIN=min(CHECKS_percent_MAX(1,:));

%% CHECKS dollars AA and EA
for aa=1:AA; CHECKS_dollars_AA(:,:,aa)=POOL_dollars_AA(:,:,aa)-( MNGR_dollars_AA(:,:,aa)+sum(CLTS_dollars_AA(:,:,1:S,aa),3) ); end
for ea=1:EA; CHECKS_dollars_EA(:,:,ea)=POOL_dollars_EA(:,:,ea)-( MNGR_dollars_EA(:,:,ea)+sum(CLTS_dollars_EA(:,:,1:S,ea),3) ); end

[CHECKS_dollars_AA_MAX(1,:,:),CHECKS_dollars_AA_MAX(2,:,:)]=max(CHECKS_dollars_AA);
 CHECKS_dollars_AA_MAX_MAX                                 =max(CHECKS_dollars_AA_MAX(1,:,:));
 CHECKS_dollars_AA_MAX_MAX_MAX                             =max(CHECKS_dollars_AA_MAX_MAX);
 
[CHECKS_dollars_AA_MIN(1,:,:),CHECKS_dollars_AA_MIN(2,:,:)]=min(CHECKS_dollars_AA);
 CHECKS_dollars_AA_MIN_MIN                                 =min(CHECKS_dollars_AA_MIN(1,:,:));
 CHECKS_dollars_AA_MIN_MIN_MIN                             =min(CHECKS_dollars_AA_MIN_MIN);
 
[CHECKS_dollars_EA_MAX(1,:,:),CHECKS_dollars_EA_MAX(2,:,:)]=max(CHECKS_dollars_EA);
 CHECKS_dollars_EA_MAX_MAX                                 =max(CHECKS_dollars_EA_MAX(1,:,:));
 CHECKS_dollars_EA_MAX_MAX_MAX                             =max(CHECKS_dollars_EA_MAX_MAX);
 
[CHECKS_dollars_EA_MIN(1,:,:),CHECKS_dollars_EA_MIN(2,:,:)]=min(CHECKS_dollars_EA);
 CHECKS_dollars_EA_MIN_MIN                                 =min(CHECKS_dollars_EA_MIN(1,:,:));
 CHECKS_dollars_EA_MIN_MIN_MIN                             =min(CHECKS_dollars_EA_MIN_MIN);
 
 %% CHECKS percent AA
for aa=1:AA; CHECKS_percent_AA(:,:,aa)= 1 - MNGR_percent_AA(:,:,aa) - sum(CLTS_percent_AA(:,:,1:S,aa),3); end
for t=1:T
    for aa=1:AA
        if CHECKS_percent_AA(t,1,aa)==1; CHECKS_percent_AA(t,1,aa)=0; end
        if CHECKS_percent_AA(t,2,aa)==1; CHECKS_percent_AA(t,2,aa)=0; end
    end
end

[CHECKS_percent_AA_MAX(1,:,:),CHECKS_percent_AA_MAX(2,:,:)]=max(CHECKS_percent_AA);
 CHECKS_percent_AA_MAX_MAX                                 =max(CHECKS_percent_AA_MAX(1,:,:));
 CHECKS_percent_AA_MAX_MAX_MAX                             =max(CHECKS_percent_AA_MAX_MAX);
 
[CHECKS_percent_AA_MIN(1,:,:),CHECKS_percent_AA_MIN(2,:,:)]=min(CHECKS_percent_AA);
 CHECKS_percent_AA_MIN_MIN                                 =min(CHECKS_percent_AA_MIN(1,:,:));
 CHECKS_percent_AA_MIN_MIN_MIN                             =min(CHECKS_percent_AA_MIN_MIN);

 
%% SAVE WORKSPACE
save('AAA_WorkSpace.mat');

%% POOL
%  EQUITY (by SERIES)
%  Base Version
import mlreportgen.dom.*;
filename=strcat('POOL_Equity_bySERIES_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_Equity.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

CONTENT{1,1}=numberFormatter(MNGR_dollars(END,end),'$###,###');
CONTENT{1,2}=numberFormatter(MNGR_percent(END,end),'##.####%');
CONTENT{1,3}=numberFormatter(MNGR_shares( END,end),'##.####');
CONTENT{1,4}=numberFormatter(MNGR_prices( END,end),'#,###.##');
CONTENT{1,5}='07-Jun-2016';

RowName{1,1}=['S00  ',MNGR_name];

for s=2:S+1
    CONTENT{s,1}=numberFormatter(CLTS_dollars(END,end,s-1),'$###,###');
    CONTENT{s,2}=numberFormatter(CLTS_percent(END,end,s-1),'##.####%');
    CONTENT{s,3}=numberFormatter(CLTS_shares( END,end,s-1),'##.####');
    CONTENT{s,4}=numberFormatter(CLTS_prices( END,end,s-1),'#,###.##');
%     CONTENT{s,5}=datestr(DATES(CLTS_dates_SUB_tinT(s-1),1),'dd-mmm-yyyy');
    CONTENT{s,5}=datestr(CLTS_dates_SUB_dtnm(s-1),'dd-mmm-yyyy');
    FormatSpec='S%02d  %s';
    RowName{s,1}=sprintf(FormatSpec,s-1,CLTS_Names{s-1});
end
    CONTENT{S+2,1}=numberFormatter(MNGR_dollars(END,end)+sum(CLTS_dollars(END,end,:),3),'$###,###');
    CONTENT{S+2,2}=numberFormatter(MNGR_percent(END,end)+sum(CLTS_percent(END,end,:),3),'##.####%');
    CONTENT{S+2,3}='...';
    CONTENT{S+2,4}='...';
    CONTENT{S+2,5}='...';
 
    RowName{S+2,1}='Total.................................';
    
TabObj = cell2table(CONTENT,'RowNames',RowName);
TabObj.Properties.VariableNames = {'EndValUSD' 'PercentEQ' 'NumberShares' 'PricePerShare' 'SubscriptionDate'};

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE';        value=TabObj;                                       append(DocObj,value);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);

clear filename template output DocObj TabObj EndValueUSD CONTENT RowName PDFfile DOCfile 


%% POOL
%  EQUITY (by SERIES)
%  CRS Version
import mlreportgen.dom.*;
filename=strcat('POOL_Equity_bySERIES_CRS_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_Equity_CRS.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

CONTENT{1,1}=numberFormatter(MNGR_dollars(END,end),'$###,###');
CONTENT{1,2}=numberFormatter(sum(MNGR_SubRed(BEG:END,4)),'$###,###');
CONTENT{1,3}=MNGR_Identifier;
CONTENT{1,4}=MNGR_TaxResidence;
CONTENT{1,5}=MNGR_Address;
CONTENT{1,6}=MNGR_TIN;
CONTENT{1,7}=MNGR_PlaceBirth;
CONTENT{1,8}=MNGR_DateBirth;
CONTENT{1,9}=MNGR_Phone;

RowName{1,1}=['S00   ',MNGR_name];

for s=2:S+1
    CONTENT{s,1}=numberFormatter(CLTS_dollars(END,end,s-1),'$###,###');
    CONTENT{s,2}=numberFormatter(sum(CLTS_SubRed(BEG:END,4,s-1)),'$###,###');
    CONTENT{s,3}=CLTS_Identifier(s-1);
    CONTENT{s,4}=CLTS_TaxResidence(s-1);
    CONTENT{s,5}=CLTS_Address(s-1);
    CONTENT{s,6}=CLTS_TIN(s-1);
    CONTENT{s,7}=CLTS_PlaceBirth(s-1);
    CONTENT{s,8}=CLTS_DateBirth(s-1);
    CONTENT{s,9}=CLTS_Phone(s-1);
    
    FormatSpec='S%02d   %s';
    RowName{s,1}=sprintf(FormatSpec,s-1,CLTS_Names{s-1});
end
    CONTENT{S+2,1}=numberFormatter(MNGR_dollars(END,end)+sum(CLTS_dollars(END,end,:),3),'$###,###');
    CONTENT{S+2,2}=numberFormatter(sum(MNGR_SubRed(BEG:END,4))+sum(sum(CLTS_SubRed(BEG:END,4,:),1),3),'$###,###');
    CONTENT{S+2,3}='...';
    CONTENT{S+2,4}='...';
    CONTENT{S+2,5}='...';
    CONTENT{S+2,6}='...';
    CONTENT{S+2,7}='...';
    CONTENT{S+2,8}='...';
    CONTENT{S+2,9}='...';
     
    RowName{S+2,1}='Total';
    
TabObj = cell2table(CONTENT,'RowNames',RowName);
TabObj.Properties.VariableNames = {'EndValUSD' 'Redemptions' 'Identifier' 'TaxResid' 'Address' 'TIN' 'PlaceBirth' 'DateBirth' 'Phone'};

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE';        value=TabObj;                                       append(DocObj,value);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);

clear filename template output DocObj TabObj EndValueUSD CONTENT RowName PDFfile DOCfile 


%% POOL
%  EQUITY by NAME
import mlreportgen.dom.*;
filename=strcat('POOL_Equity_byNAME_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_Equity.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

% CLIENTS
[CLTS_Names_unique,~,~] = unique(CLTS_Names);
                N_clients  = length(CLTS_Names_unique);
                
for n=1:N_clients
    STRFIND=strfind(CLTS_Names,CLTS_Names_unique{n});
    I=find(~cellfun(@isempty,STRFIND));
    % sum over SameName
    CLTS_dollars_sum_byNAME(n,1)=sum(CLTS_dollars(END,end,I),3);
    CLTS_percent_sum_byNAME(n,1)=sum(CLTS_percent(END,end,I),3);
end
    % sort
    [CLTS_dollars_sum_byNAME_sort, I_sort]=sort(CLTS_dollars_sum_byNAME,'descend');
     CLTS_percent_sum_byNAME_sort = CLTS_percent_sum_byNAME(I_sort);
           CLTS_Names_unique_sort =       CLTS_Names_unique(I_sort);
    
CONTENT{1,1}=numberFormatter(MNGR_dollars(END,end),'$###,###');
CONTENT{1,2}=numberFormatter(MNGR_percent(END,end),'##.##%');

RowName{1,1}=[num2str(1,'%02d'),'    ',MNGR_name];

for n=2:N_clients+1
    CONTENT{n,1}=numberFormatter(CLTS_dollars_sum_byNAME_sort(n-1),'$###,###');
    CONTENT{n,2}=numberFormatter(CLTS_percent_sum_byNAME_sort(n-1),'##.##%');
    RowName{n,1}      =[num2str(n,'%02d'),'    ',CLTS_Names_unique_sort{n-1}];
end
    CONTENT{N_clients+2,1}=numberFormatter(MNGR_dollars(END,end)+sum(CLTS_dollars_sum_byNAME_sort),'$###,###');
    CONTENT{N_clients+2,2}=numberFormatter(MNGR_percent(END,end)+sum(CLTS_percent_sum_byNAME_sort),'##.##%');
 
    RowName{N_clients+2,1}='Total.............................................................';
    
TabObj = cell2table(CONTENT,'RowNames',RowName);
TabObj.Properties.VariableNames = {'EndValUSD' 'PercentEQ'};

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE';        value=TabObj;                                       append(DocObj,value);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);

clear filename template output DocObj TabObj EndValueUSD CONTENT RowName PDFfile DOCfile
clear N_clients

%% POOL
%  Dollar Ledger Summary
import mlreportgen.dom.*;
filename=strcat('POOL_DollarLedgerSummary_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_DollarLedgerSummary.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'BEG';         value=cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY'));
              case 'END';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));
              
              case 'BAL_beg';     value=numberFormatter(        POOL_dollars(BEG    ,01)   ,'$###,###');    % BEG
              case 'SUBS';        value=numberFormatter(sum(    POOL_dollars(BEG:END,02),1),'$###,###');    % SUBSCRIPTIONS
              
              case 'INCOME_tot';  value=numberFormatter(sum(sum(POOL_dollars(BEG:END,[05 10 14]),2),1)                     ,'$###,###'); % INCOME  BAincome + AAaccrual + AAwriteoff
              case 'CAP_GAIN';    value=numberFormatter(sum(sum(POOL_dollars(BEG:END,[06 07 08 09 11 12 13 15 17 19]),2),1),'$###,###'); % CAP.GN  SubCostADDback + BA_gain + DEPOSITS + PAYOUTS + AA_gain + AA_fx + AA_reversal + AA_writeoff_gain + EA_cptlztn + RedCostADDback)
             
              case 'COSTS_tot';   value=numberFormatter(sum(sum(POOL_dollars(BEG:END,[03 16 18]),2),1),'$###,###'); % COSTS: total
              case 'COSTS_subs';  value=numberFormatter(sum(    POOL_dollars(BEG:END,03)           ,1),'$###,###'); % SubCosts
              case 'COSTS_admin'; value=numberFormatter(sum(    POOL_dollars(BEG:END,16)           ,1),'$###,###'); % AdminCosts (EA_depreciation)
              case 'COSTS_reds';  value=numberFormatter(sum(    POOL_dollars(BEG:END,18)           ,1),'$###,###'); % RedCosts
              
              case 'REDS';        value=numberFormatter(sum(    POOL_dollars(BEG:END,26)           ,1),'$###,###'); % REDEMPTIONS
              case 'BAL_end';     value=numberFormatter(        POOL_dollars(    END,end)             ,'$###,###'); % END
                  
              case 'COSTS_setup';      value=numberFormatter(sum(    MNGR_dollars(BEG:END,21)        ,1),'$###,###'); % SetupCosts
              case 'FEE_tot';          value=numberFormatter(sum(sum(MNGR_dollars(BEG:END,[22 23]),2),1),'$###,###'); % MgmtFee + PerfFee   (net of finders)
              case 'FEE_mgmt';         value=numberFormatter(sum(    MNGR_dollars(BEG:END, 22)       ,1),'$###,###'); % MgmtFee             (net of finders)
              case 'FEE_perf';         value=numberFormatter(sum(    MNGR_dollars(BEG:END,    23)    ,1),'$###,###'); % PerfFee             (net of finders)
                  
              case 'FEE_mgmt_GROSS';   value=numberFormatter(sum( MNGR_MgmtFee(BEG:END,1) ,1),'$###,###'); % FEE:    Management  (gross)
              case 'FF_mgmt';          value=numberFormatter(sum( MNGR_MgmtFee(BEG:END,2) ,1),'$###,###'); % Finder: Management    
              case 'FEE_mgmt_NET';     value=numberFormatter(sum( MNGR_MgmtFee(BEG:END,3) ,1),'$###,###'); % FEE:    Management  (net)
                  
              case 'FEE_perf_GROSS';   value=numberFormatter(sum( MNGR_PerfFee(BEG:END,1) ,1),'$###,###'); % FEE:    Performance  (gross)
              case 'FF_perf';          value=numberFormatter(sum( MNGR_PerfFee(BEG:END,2) ,1),'$###,###'); % Finder: Performance    
              case 'FEE_perf_NET';     value=numberFormatter(sum( MNGR_PerfFee(BEG:END,3) ,1),'$###,###'); % FEE:    Performance  (net)
          end
          append(DocObj,value{:});
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj PDFfile DOCfile


%%  MONTHLY PERFORMANCE TABLE
%   POOL shadow SHARES
import mlreportgen.dom.*;
filename = strcat('Performance_Table_',datestr(DATES(END,1),'DDmmmYYYY'));
template = strcat('reports\templates\','POOL_PerformanceTable.dotx');
output   = strcat('reports\',filename);
DocObj   = Document(output,'docx',template);

% all MONTHS
for m=1:M
    POOL_return_M(m,1)=DATES(EoM(m),1);
    
    if m==1; temp=cumprod(1+POOL_return(         1:EoM(m)))-1; end
    if m>=2; temp=cumprod(1+POOL_return(EoM(m-1)+1:EoM(m)))-1; end
    
    POOL_return_M(m,2)=temp(end);
    clear temp
end
LM=month(DATES(end,1))+1;

TABLE{1,1}='2015'; TABLE(1,2:13)=numberFormatter(POOL_return_M(01:12 ,2)','#0.0%'); YTD=cumprod(1+POOL_return_M(01:12 ,2))-1; TABLE(1,14)=numberFormatter(YTD(end),'#0.0%');
TABLE{2,1}='2016'; TABLE(2,2:13)=numberFormatter(POOL_return_M(13:24 ,2)','#0.0%'); YTD=cumprod(1+POOL_return_M(13:24 ,2))-1; TABLE(2,14)=numberFormatter(YTD(end),'#0.0%');
TABLE{3,1}='2017'; TABLE(3,2:13)=numberFormatter(POOL_return_M(25:36 ,2)','#0.0%'); YTD=cumprod(1+POOL_return_M(25:36 ,2))-1; TABLE(3,14)=numberFormatter(YTD(end),'#0.0%');
TABLE{4,1}='2018'; TABLE(4,2:13)=numberFormatter(POOL_return_M(37:48 ,2)','#0.0%'); YTD=cumprod(1+POOL_return_M(37:48 ,2))-1; TABLE(4,14)=numberFormatter(YTD(end),'#0.0%');
TABLE{5,1}='2019'; TABLE(5,2:13)=numberFormatter(POOL_return_M(49:60 ,2)','#0.0%'); YTD=cumprod(1+POOL_return_M(49:60 ,2))-1; TABLE(5,14)=numberFormatter(YTD(end),'#0.0%');
TABLE{6,1}='2020'; TABLE(6,2:13)=numberFormatter(POOL_return_M(61:72 ,2)','#0.0%'); YTD=cumprod(1+POOL_return_M(61:72 ,2))-1; TABLE(6,14)=numberFormatter(YTD(end),'#0.0%');
TABLE{7,1}='2021'; TABLE(7,2:13)=numberFormatter(POOL_return_M(73:84 ,2)','#0.0%'); YTD=cumprod(1+POOL_return_M(73:84 ,2))-1; TABLE(7,14)=numberFormatter(YTD(end),'#0.0%');
TABLE{8,1}='2022'; TABLE(8,2:LM)=numberFormatter(POOL_return_M(85:end,2)','#0.0%'); YTD=cumprod(1+POOL_return_M(85:end,2))-1; TABLE(8,14)=numberFormatter(YTD(end),'#0.0%');

TABLE(8,LM+1:13)=cellstr('--');

TABLE  =flipud(TABLE);

TABLE=[{'' 'Jan' 'Feb' 'Mar' 'Apr' 'May' 'Jun' 'Jul' 'Aug' 'Sep' 'Oct' 'Nov' 'Dec' 'YTD'};TABLE];

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';    value=datestr(DATES(END,1),'DD-mmm-YYYY'); append(DocObj,value);
              case 'TABLE';   value=TABLE;                               append(DocObj,value);
          end
          moveToNextHole(DocObj);
end


close(DocObj);
clear filename template output PDFfile DOCfile DocObj
clear TABLE EndValueUSD RowName


%% AUDIT
%  Balance Sheet
if DATES_datevec(BEG,2)==1 && DATES_datevec(BEG,3)==1 && DATES_datevec(END,2)==12 && DATES_datevec(END,3)==31 && DATES_datevec(BEG,1)==DATES_datevec(END,1)

    import mlreportgen.dom.*;
    filename=strcat('AUDIT_BalanceSheet_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'));
    template=strcat('reports\templates\','AUDIT_BalanceSheet.dotx');
    output=strcat('reports\',filename);
    DocObj = Document(output,'docx',template);

                % CURRENT period END
                    END_c    =END;
                    END_c_txt=cellstr(datestr(DATES(END_c,1),'DD-mmm-YYYY'));
                
                % PRVIOUS period END
                    if DATES_datevec(END_c,1)>=2016
                    END_p_dtvc=DATES_datevec(END_c,:)-[1,0,0,0,0,0];
                    END_p_txt =cellstr(datestr(END_p_dtvc,'DD-mmm-YYYY'));
                    END_p     = find(DATES==datenum(END_p_txt,'DD-mmm-YYYY'));
                    end
                    
                % Single out CASH
                    IND_CASH=strfind(POOL_types_BA,'CASH');
                    IND_CASH=find(~cellfun(@isempty,IND_CASH));
                    
                % CASH and Held Custody
                    AUDIT_BalanceSheet_CASH_c = sum(POOL_dollars_BA(END_c,end,IND_CASH),3);
 if exist('END_p'); AUDIT_BalanceSheet_CASH_p = sum(POOL_dollars_BA(END_p,end,IND_CASH),3); end
             
                    AUDIT_BalanceSheet_HeldCust_c =-(POOL_dollars_HC_sub(END_c,end)+POOL_dollars_HC_red(END_c,end));
 if exist('END_p'); AUDIT_BalanceSheet_HeldCust_p =-(POOL_dollars_HC_sub(END_p,end)+POOL_dollars_HC_red(END_p,end)); end
 
                % PREPAID expenses and ACCRUED income
                    AUDIT_BalanceSheet_AA_asset_c = reshape(POOL_dollars_AA(END_c,end,:),[AA,1]); AUDIT_BalanceSheet_AA_asset_c = sum(AUDIT_BalanceSheet_AA_asset_c(find(AUDIT_BalanceSheet_AA_asset_c>0)));
 if exist('END_p'); AUDIT_BalanceSheet_AA_asset_p = reshape(POOL_dollars_AA(END_p,end,:),[AA,1]); AUDIT_BalanceSheet_AA_asset_p = sum(AUDIT_BalanceSheet_AA_asset_p(find(AUDIT_BalanceSheet_AA_asset_p>0))); end
                    
                    AUDIT_BalanceSheet_EA_asset_c = reshape(POOL_dollars_EA(END_c,end,:),[EA,1]); AUDIT_BalanceSheet_EA_asset_c = sum(AUDIT_BalanceSheet_EA_asset_c(find(AUDIT_BalanceSheet_EA_asset_c>0)));
 if exist('END_p'); AUDIT_BalanceSheet_EA_asset_p = reshape(POOL_dollars_EA(END_p,end,:),[EA,1]); AUDIT_BalanceSheet_EA_asset_p = sum(AUDIT_BalanceSheet_EA_asset_p(find(AUDIT_BalanceSheet_EA_asset_p>0))); end
                    
                    AUDIT_BalanceSheet_PrpdAcrd_c = AUDIT_BalanceSheet_AA_asset_c + AUDIT_BalanceSheet_EA_asset_c;
 if exist('END_p'); AUDIT_BalanceSheet_PrpdAcrd_p = AUDIT_BalanceSheet_AA_asset_p + AUDIT_BalanceSheet_EA_asset_p; end
                
                % CURRENT ASSETS
                    AUDIT_BalanceSheet_CrntAsst_c = AUDIT_BalanceSheet_CASH_c + AUDIT_BalanceSheet_PrpdAcrd_c;
 if exist('END_p'); AUDIT_BalanceSheet_CrntAsst_p = AUDIT_BalanceSheet_CASH_p + AUDIT_BalanceSheet_PrpdAcrd_p; end
                
                % FINANCIAL ASSETS
                    AUDIT_BalanceSheet_FincAsst_c = sum(POOL_dollars_BA(END_c,end,:),3) - AUDIT_BalanceSheet_CASH_c;
 if exist('END_p'); AUDIT_BalanceSheet_FincAsst_p = sum(POOL_dollars_BA(END_p,end,:),3) - AUDIT_BalanceSheet_CASH_p; end
 
                % NON-CURRENT ASSETS
                    AUDIT_BalanceSheet_nCrrAsst_c = AUDIT_BalanceSheet_FincAsst_c;
 if exist('END_p'); AUDIT_BalanceSheet_nCrrAsst_p = AUDIT_BalanceSheet_FincAsst_p; end
                    
                % TOTAL ASSETS
                    AUDIT_BalanceSheet_TotlAsst_c = AUDIT_BalanceSheet_CrntAsst_c + AUDIT_BalanceSheet_nCrrAsst_c;
 if exist('END_p'); AUDIT_BalanceSheet_TotlAsst_p = AUDIT_BalanceSheet_CrntAsst_p + AUDIT_BalanceSheet_nCrrAsst_p; end
                
                % ACCRUED expenses and DEFERRED Income
                    AUDIT_BalanceSheet_AA_lblty_c = reshape(POOL_dollars_AA(END_c,end,:),[AA,1]); AUDIT_BalanceSheet_AA_lblty_c = sum(AUDIT_BalanceSheet_AA_lblty_c(find(AUDIT_BalanceSheet_AA_lblty_c<0)));
 if exist('END_p'); AUDIT_BalanceSheet_AA_lblty_p = reshape(POOL_dollars_AA(END_p,end,:),[AA,1]); AUDIT_BalanceSheet_AA_lblty_p = sum(AUDIT_BalanceSheet_AA_lblty_p(find(AUDIT_BalanceSheet_AA_lblty_p<0))); end
 
                    AUDIT_BalanceSheet_EA_lblty_c = reshape(POOL_dollars_EA(END_c,end,:),[EA,1]); AUDIT_BalanceSheet_EA_lblty_c = sum(AUDIT_BalanceSheet_EA_lblty_c(find(AUDIT_BalanceSheet_EA_lblty_c<0)));
 if exist('END_p'); AUDIT_BalanceSheet_EA_lblty_p = reshape(POOL_dollars_EA(END_p,end,:),[EA,1]); AUDIT_BalanceSheet_EA_lblty_p = sum(AUDIT_BalanceSheet_EA_lblty_p(find(AUDIT_BalanceSheet_EA_lblty_p<0))); end
 
                    AUDIT_BalanceSheet_AcrdDfrd_c =-( AUDIT_BalanceSheet_AA_lblty_c + AUDIT_BalanceSheet_EA_lblty_c );
 if exist('END_p'); AUDIT_BalanceSheet_AcrdDfrd_p =-( AUDIT_BalanceSheet_AA_lblty_p + AUDIT_BalanceSheet_EA_lblty_p ); end
 
                % CURRENT LIABILITES
                    AUDIT_BalanceSheet_CrntLbly_c = AUDIT_BalanceSheet_AcrdDfrd_c + AUDIT_BalanceSheet_HeldCust_c;
 if exist('END_p'); AUDIT_BalanceSheet_CrntLbly_p = AUDIT_BalanceSheet_AcrdDfrd_p + AUDIT_BalanceSheet_HeldCust_p; end
 
                % TOTAL LIABILITIES
                    AUDIT_BalanceSheet_TotlLbly_c = AUDIT_BalanceSheet_CrntLbly_c;
 if exist('END_p'); AUDIT_BalanceSheet_TotlLbly_p = AUDIT_BalanceSheet_CrntLbly_p; end
 
                % CONTRIBUTED CAPITAL
                    AUDIT_BalanceSheet_ContCapt_c = sum(POOL_SubRed(1:END_c,1)) + sum(POOL_SubRed(1:END_c,4));
 if exist('END_p'); AUDIT_BalanceSheet_ContCapt_p = sum(POOL_SubRed(1:END_p,1)) + sum(POOL_SubRed(1:END_p,4)); end
 
                % SHAREHOLDERS' EQUITY
                    AUDIT_BalanceSheet_ShrhEqty_c = POOL_dollars(END_c,end);
 if exist('END_p'); AUDIT_BalanceSheet_ShrhEqty_p = POOL_dollars(END_p,end); end
 
                % CUMULATIVE EARNING
                    AUDIT_BalanceSheet_CumlEarn_c = AUDIT_BalanceSheet_ShrhEqty_c - AUDIT_BalanceSheet_ContCapt_c;
 if exist('END_p'); AUDIT_BalanceSheet_CumlEarn_p = AUDIT_BalanceSheet_ShrhEqty_p - AUDIT_BalanceSheet_ContCapt_p; end
     
                % LIABILITIES + EQUITY
                    AUDIT_BalanceSheet_LblyEqty_c = AUDIT_BalanceSheet_TotlLbly_c + AUDIT_BalanceSheet_ShrhEqty_c;
 if exist('END_p'); AUDIT_BalanceSheet_LblyEqty_p = AUDIT_BalanceSheet_TotlLbly_p + AUDIT_BalanceSheet_ShrhEqty_p; end
                   
    moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'END_c';         value=END_c_txt;
              case 'END_p';         value=END_p_txt;
              
              case 'CASH_c';        value=numberFormatter(AUDIT_BalanceSheet_CASH_c,'$###,###');
              case 'CASH_p';        value=numberFormatter(AUDIT_BalanceSheet_CASH_p,'$###,###');
                  
              case 'PrpdAcrd_c';    value=numberFormatter(AUDIT_BalanceSheet_PrpdAcrd_c,'$###,###');
              case 'PrpdAcrd_p';    value=numberFormatter(AUDIT_BalanceSheet_PrpdAcrd_p,'$###,###');
              
              case 'CrntAsst_c';    value=numberFormatter(AUDIT_BalanceSheet_CrntAsst_c,'$###,###');
              case 'CrntAsst_p';    value=numberFormatter(AUDIT_BalanceSheet_CrntAsst_p,'$###,###');
                  
              case 'FincAsst_c';    value=numberFormatter(AUDIT_BalanceSheet_FincAsst_c,'$###,###');
              case 'FincAsst_p';    value=numberFormatter(AUDIT_BalanceSheet_FincAsst_p,'$###,###');
                  
              case 'nCrrAsst_c';    value=numberFormatter(AUDIT_BalanceSheet_nCrrAsst_c,'$###,###');
              case 'nCrrAsst_p';    value=numberFormatter(AUDIT_BalanceSheet_nCrrAsst_p,'$###,###');
                  
              case 'TotlAsst_c';    value=numberFormatter(AUDIT_BalanceSheet_TotlAsst_c,'$###,###');
              case 'TotlAsst_p';    value=numberFormatter(AUDIT_BalanceSheet_TotlAsst_p,'$###,###');
                  
              case 'HeldCust_c';    value=numberFormatter(AUDIT_BalanceSheet_HeldCust_c,'$###,###');
              case 'HeldCust_p';    value=numberFormatter(AUDIT_BalanceSheet_HeldCust_p,'$###,###');
                  
              case 'AcrdDfrd_c';    value=numberFormatter(AUDIT_BalanceSheet_AcrdDfrd_c,'$###,###');
              case 'AcrdDfrd_p';    value=numberFormatter(AUDIT_BalanceSheet_AcrdDfrd_p,'$###,###');
                  
              case 'CrntLbly_c';    value=numberFormatter(AUDIT_BalanceSheet_CrntLbly_c,'$###,###');
              case 'CrntLbly_p';    value=numberFormatter(AUDIT_BalanceSheet_CrntLbly_p,'$###,###');
                  
              case 'TotlLbly_c';    value=numberFormatter(AUDIT_BalanceSheet_TotlLbly_c,'$###,###');
              case 'TotlLbly_p';    value=numberFormatter(AUDIT_BalanceSheet_TotlLbly_p,'$###,###');
              
              case 'ContCapt_c';    value=numberFormatter(AUDIT_BalanceSheet_ContCapt_c,'$###,###');
              case 'ContCapt_p';    value=numberFormatter(AUDIT_BalanceSheet_ContCapt_p,'$###,###');
              
              case 'CumlEarn_c';    value=numberFormatter(AUDIT_BalanceSheet_CumlEarn_c,'$###,###');
              case 'CumlEarn_p';    value=numberFormatter(AUDIT_BalanceSheet_CumlEarn_p,'$###,###');
                  
              case 'ShrhEqty_c';    value=numberFormatter(AUDIT_BalanceSheet_ShrhEqty_c,'$###,###');
              case 'ShrhEqty_p';    value=numberFormatter(AUDIT_BalanceSheet_ShrhEqty_p,'$###,###');
                  
              case 'LblyEqty_c';    value=numberFormatter(AUDIT_BalanceSheet_LblyEqty_c,'$###,###');
              case 'LblyEqty_p';    value=numberFormatter(AUDIT_BalanceSheet_LblyEqty_p,'$###,###');
          end
          append(DocObj,value{:});
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj PDFfile DOCfile

end

%% POOL
%  BALANCE SHEET by BA,AA,EA,HC
import mlreportgen.dom.*;
filename=strcat('POOL_BalanceSheet_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_BalanceSheet.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

% BA
[POOL_dollars_EndVal_BA_sort,I_BA_EndVal_sort] = sort(reshape(POOL_dollars_BA(END,end,:),[BA,1]),'descend');
 POOL_percent_EndVal_BA_sort                   =      reshape(POOL_percent_BA(END,end,:),[BA,1]);

 POOL_percent_EndVal_BA_sort = POOL_percent_EndVal_BA_sort(I_BA_EndVal_sort);
 POOL_ccy_BA_sort            = POOL_ccy_BA(                I_BA_EndVal_sort);
 POOL_names_BA_sort          = POOL_names_BA(              I_BA_EndVal_sort);
 POOL_types_BA_sort          = POOL_types_BA(              I_BA_EndVal_sort);
 POOL_IDs_BA_sort            = POOL_IDs_BA(                I_BA_EndVal_sort);
 POOL_title_BA_sort          = POOL_title_BA(              I_BA_EndVal_sort);

for ba=1:BA
    CONTENT{ba,6}=numberFormatter(POOL_dollars_EndVal_BA_sort(ba),'$###,###');
    CONTENT{ba,5}=numberFormatter(POOL_percent_EndVal_BA_sort(ba),'##.#%');
    CONTENT{ba,4}=POOL_title_BA_sort{ba};
    CONTENT{ba,3}=POOL_IDs_BA_sort{  ba};
    CONTENT{ba,2}=POOL_ccy_BA_sort{  ba};
    CONTENT{ba,1}=POOL_types_BA_sort{ba};
    RowName{ba}=sprintf('(BA%02d) %s',ba,POOL_names_BA_sort{ba});
end
    CONTENT{BA+1,6}=numberFormatter(sum(POOL_dollars_EndVal_BA_sort),'$###,###');
    CONTENT{BA+1,5}=numberFormatter(sum(POOL_percent_EndVal_BA_sort),'##.#%');
    CONTENT{BA+1,4}='---';
    CONTENT{BA+1,3}='---';
    CONTENT{BA+1,2}='---';
    CONTENT{BA+1,1}='---';
    RowName{BA+1}='(BA99) Total BA';
    
TabObj_BA = cell2table(CONTENT,'RowNames',RowName);
TabObj_BA.Properties.VariableNames = {'AssetType' 'CCY' 'AccountID' 'AccountHolder' 'PercentEQ' 'EndValUSD'};
clear CONTENT RowName

% AA
[POOL_dollars_EndVal_AA_sort,I_AA_EndVal_sort]=sort(reshape(POOL_dollars_AA(END,end,:),[AA,1]),'descend');
 POOL_percent_EndVal_AA_sort=reshape(POOL_percent_AA(END,end,:),[AA,1]);

 POOL_percent_EndVal_AA_sort = POOL_percent_EndVal_AA_sort(I_AA_EndVal_sort);
 POOL_ccy_AA_sort            = POOL_ccy_AA(                I_AA_EndVal_sort);
 POOL_names_AA_sort          = POOL_names_AA(              I_AA_EndVal_sort);
 POOL_types_AA_sort          = POOL_types_AA(              I_AA_EndVal_sort);
 POOL_IDs_AA_sort            = POOL_IDs_AA(                I_AA_EndVal_sort);
 POOL_title_AA_sort          = POOL_title_AA(              I_AA_EndVal_sort);

for aa=1:AA
    CONTENT{aa,5}=numberFormatter(POOL_dollars_EndVal_AA_sort(aa),'$###,###');
    CONTENT{aa,4}=numberFormatter(POOL_percent_EndVal_AA_sort(aa),'##.#%');
    CONTENT{aa,3}=POOL_IDs_AA_sort{aa};
    CONTENT{aa,2}=POOL_ccy_AA_sort{  aa};
    CONTENT{aa,1}=POOL_types_AA_sort{aa};
    RowName{aa,1}=sprintf('(AA%02d) %s',aa,POOL_names_AA_sort{aa});
end
    CONTENT{AA+1,5}=numberFormatter(sum(POOL_dollars_EndVal_AA_sort),'$###,###');
    CONTENT{AA+1,4}=numberFormatter(sum(POOL_percent_EndVal_AA_sort),'##.#%');
    CONTENT{AA+1,3}='---';
    CONTENT{AA+1,2}='---';
    CONTENT{AA+1,1}='---';
    RowName{AA+1,1}='(AA99) Total AA';
    
TabObj_AA = cell2table(CONTENT,'RowNames',RowName);
TabObj_AA.Properties.VariableNames = {'AssetType' 'CCY' 'AccountNumber' 'PercentEQ' 'EndValUSD'};
clear CONTENT RowName

% EA
[POOL_dollars_EndVal_EA_sort,I_EA_EndVal_sort] = sort(reshape(POOL_dollars_EA(END,end,:),[EA,1]),'descend');
 POOL_percent_EndVal_EA_sort                   = sort(reshape(POOL_percent_EA(END,end,:),[EA,1]),'descend');
          POOL_names_EA_sort                                 =POOL_names_EA(I_EA_EndVal_sort);
          POOL_types_EA_sort                                 =POOL_types_EA(I_EA_EndVal_sort);

for ea=1:EA
    CONTENT{ea,3}=numberFormatter(POOL_dollars_EndVal_EA_sort(ea),'$###,###');
    CONTENT{ea,2}=numberFormatter(POOL_percent_EndVal_EA_sort(ea),'##.#%');
    CONTENT{ea,1}=POOL_names_EA_sort{ea};
    RowName{ea,1}=sprintf('(EA%02d) %s',ea,POOL_types_EA_sort{ea});
end
    CONTENT{EA+1,3}=numberFormatter(sum(POOL_dollars_EndVal_EA_sort ),'$###,###');
    CONTENT{EA+1,2}=numberFormatter(sum(POOL_percent_EndVal_EA_sort),'##.#%');
    CONTENT{EA+1,1}='---';
    RowName{EA+1,1}='(EA99) Total EA';
        
    TabObj_EA = cell2table(CONTENT,'RowNames',RowName);
    TabObj_EA.Properties.VariableNames = {'ServiceProvider' 'PercentEQ' 'EndValUSD'};
    clear CONTENT RowName

% HC
    CONTENT{1,2} = numberFormatter( POOL_dollars_HC_sub(END,end) ,'$###,###');
    CONTENT{1,1} = numberFormatter( POOL_percent_HC_sub(END,end) ,'##.#%');
    RowName{1,1} = 'HELD in CUSTODAY: SUBSCRIPTIONS';
    
    CONTENT{2,2} = numberFormatter( POOL_dollars_HC_red(END,end) ,'$###,###');
    CONTENT{2,1} = numberFormatter( POOL_percent_HC_red(END,end) ,'##.#%');
    RowName{2,1} = 'HELD in CUSTODAY: REDEMPTIONS';
    
    CONTENT{3,2} = numberFormatter( POOL_dollars_HC_sub(END,end) + POOL_dollars_HC_red(END,end) ,'$###,###');
    CONTENT{3,1} = numberFormatter( POOL_percent_HC_sub(END,end) + POOL_percent_HC_red(END,end) ,'##.#%');
    RowName{3,1} = 'Total';

    TabObj_HC = cell2table(CONTENT,'RowNames',RowName);
    TabObj_HC.Properties.VariableNames = {'PercentEQ' 'EndValUSD'};
    clear CONTENT RowName
    
% TOTAL
CONTENT{1,2}=numberFormatter(POOL_dollars(END,end),'$###,###');
CONTENT{1,1}=numberFormatter(   sum(POOL_percent_EndVal_BA_sort) ...
                              + sum(POOL_percent_EndVal_AA_sort) ...
                              + sum(POOL_percent_EndVal_EA_sort) ...
                              + POOL_percent_HC_sub(END,end)     ...
                              + POOL_percent_HC_red(END,end)     ,'###.##%');
                          
RowName{1,1}='FUND NAV';
TabObj_NAV = cell2table(CONTENT,'RowNames',RowName);
TabObj_NAV.Properties.VariableNames = {'PercentEQ' 'EndValUSD'};
clear CONTENT RowName

% PLACE in HOLES
moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));         append(DocObj,value{:});
              case 'TABLE_BA';     value=TabObj_BA;                                            append(DocObj,value);
              case 'TABLE_AA';     value=TabObj_AA;                                            append(DocObj,value);
              case 'TABLE_EA';     value=TabObj_EA;                                            append(DocObj,value);
              case 'TABLE_HC';     value=TabObj_HC;                                            append(DocObj,value);
              case 'POOL_NAV';     value=TabObj_NAV;                                           append(DocObj,value);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);

clear filename template output DocObj CONTENT RowName PDFfile DOCfile
clear TabObj_BA TabObj_AA TabObj_EA TabObj_HC TabObj_NAV EndValueUSD
clear POOL_names_BA_sort ...
      POOL_names_AA_sort ...
      POOL_names_EA_sort ...
      POOL_ccy_BA_sort ...
      POOL_ccy_AA_sort ...
      POOL_ccy_EA_sort ...
      POOL_IDs_BA_sort ...
      POOL_IDs_AA_sort ...
      POOL_IDs_EA_sort ...
      POOL_title_BA_sort ...
      POOL_title_AA_sort ...
      POOL_title_EA_sort ...
      POOL_types_BA_sort ...
      POOL_types_AA_sort ...
      POOL_types_EA_sort ...
      POOL_dollars_EndVal_BA_sort ...
      POOL_dollars_EndVal_AA_sort ...
      POOL_dollars_EndVal_EA_sort ...
      POOL_percent_EndVal_BA_sort ...
      POOL_percent_EndVal_AA_sort ...
      POOL_percent_EndVal_EA_sort ...
      I_BA_EndVal_sort ...
      I_AA_EndVal_sort ...
      I_EA_EndVal_sort
      
%% POOL
%  BALANCE SHEET by CCY
import mlreportgen.dom.*;
filename=strcat('POOL_BalanceSheet_CCY_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_BalanceSheet_CCY.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

for i=1:CC
    STRFIND_BA=strcmp(POOL_ccy_BA,FX_ccy{i}); I_BA=find(STRFIND_BA==1);
    STRFIND_AA=strcmp(POOL_ccy_AA,FX_ccy{i}); I_AA=find(STRFIND_AA==1);
    
    % BA
    POOL_dollars_BA_sum_CCY(i) = sum(POOL_dollars_BA(END,end,I_BA),3);
    POOL_percent_BA_sum_CCY(i) = sum(POOL_percent_BA(END,end,I_BA),3);
    POOL_local_BA_sum_CCY(i)   =     POOL_dollars_BA_sum_CCY(i) / FX(END,i);
    
    % AA
    POOL_dollars_AA_sum_CCY(i) = sum(POOL_dollars_AA(END,end,I_AA),3);
    POOL_percent_AA_sum_CCY(i) = sum(POOL_percent_AA(END,end,I_AA),3);
    POOL_local_AA_sum_CCY(i)   =     POOL_dollars_AA_sum_CCY(i) / FX(END,i);
    
    % EA (all EAs are in USD)
    if  strcmp(FX_ccy{i},'USD')
    POOL_dollars_EA_sum_CCY(i) =sum(POOL_dollars_EA(END,end,:),3);
    POOL_percent_EA_sum_CCY(i) =sum(POOL_percent_EA(END,end,:),3);
    POOL_local_EA_sum_CCY(i)   =    POOL_dollars_EA_sum_CCY(i) / FX(END,i);
    end
    if ~strcmp(FX_ccy{i},'USD')
    POOL_dollars_EA_sum_CCY(i)=0;
    POOL_percent_EA_sum_CCY(i)=0;
    POOL_local_EA_sum_CCY(i)  =0;
    end
       
    % HC (all HCs are in USD)
    if  strcmp(FX_ccy{i},'USD')
    POOL_dollars_HC_sum_CCY(i) = POOL_dollars_HC_sub(END,end) + POOL_dollars_HC_red(END,end);
    POOL_percent_HC_sum_CCY(i) = POOL_percent_HC_sub(END,end) + POOL_percent_HC_red(END,end);
    POOL_local_HC_sum_CCY(i)   = POOL_dollars_HC_sum_CCY(i) / FX(END,i);
    end
    if ~strcmp(FX_ccy{i},'USD')
    POOL_dollars_HC_sum_CCY(i) = 0;
    POOL_percent_HC_sum_CCY(i) = 0;
    POOL_local_HC_sum_CCY(i)   = 0;
    end
    
    % ISNAN
    ISNAN_BA=isnan(POOL_local_BA_sum_CCY); POOL_local_BA_sum_CCY(1,ISNAN_BA)=0;
    ISNAN_AA=isnan(POOL_local_AA_sum_CCY); POOL_local_AA_sum_CCY(1,ISNAN_AA)=0;
    ISNAN_EA=isnan(POOL_local_EA_sum_CCY); POOL_local_EA_sum_CCY(1,ISNAN_EA)=0;
    ISNAN_HC=isnan(POOL_local_HC_sum_CCY); POOL_local_HC_sum_CCY(1,ISNAN_HC)=0;
    
    % AGGREGATE across BA,AA,EA,HC
    POOL_dollars_sum_CCY(i) =  POOL_dollars_BA_sum_CCY(i) + POOL_dollars_AA_sum_CCY(i) + POOL_dollars_EA_sum_CCY(i) + POOL_dollars_HC_sum_CCY(i);
    POOL_percent_sum_CCY(i) =  POOL_percent_BA_sum_CCY(i) + POOL_percent_AA_sum_CCY(i) + POOL_percent_EA_sum_CCY(i) + POOL_percent_HC_sum_CCY(i);
      POOL_local_sum_CCY(i) =    POOL_local_BA_sum_CCY(i) +   POOL_local_AA_sum_CCY(i) +   POOL_local_EA_sum_CCY(i) +   POOL_local_HC_sum_CCY(i);
end

% RESHAPE stacked variables
POOL_dollars_sum_CCY=reshape(POOL_dollars_sum_CCY,[CC 1]);
POOL_percent_sum_CCY=reshape(POOL_percent_sum_CCY,[CC 1]);
POOL_local_sum_CCY  =reshape(  POOL_local_sum_CCY,[CC 1]);

% SORT stacked variables
[POOL_dollars_sum_CCY_sort,I_sort]=sort(POOL_dollars_sum_CCY,'descend');
               FX_ccy_sort=              FX_ccy(I_sort)';
 POOL_percent_sum_CCY_sort=POOL_percent_sum_CCY(I_sort) ;
   POOL_local_sum_CCY_sort=  POOL_local_sum_CCY(I_sort) ;

for cc=1:CC
    CONTENT{cc,3}=numberFormatter(POOL_dollars_sum_CCY_sort(cc),'$###,###');
    CONTENT{cc,2}=numberFormatter(  POOL_local_sum_CCY_sort(cc),'###,###');
    CONTENT{cc,1}=numberFormatter(POOL_percent_sum_CCY_sort(cc),'##.#%');
    RowName{cc}=FX_ccy_sort{cc};
end
    CONTENT{CC+1,3}=numberFormatter(sum( POOL_dollars_sum_CCY_sort),'$###,###');
    CONTENT{CC+1,2}={'---'};
    CONTENT{CC+1,1}=numberFormatter(sum(POOL_percent_sum_CCY_sort),'##.#%');
    RowName{CC+1}='Total.............................................';
TabObj = cell2table(CONTENT,'RowNames',RowName);
TabObj.Properties.VariableNames = {'percent' 'LocalCCY' 'USDval'};

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE';        value=TabObj;                                       append(DocObj, TabObj);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj CONTENT TabObj RowName PDFfile DOCfile
clear ISNAN_BA ISNAN_AA ISNAN_EA
clear POOL_dollars_BA_sum_CCY ...
      POOL_dollars_AA_sum_CCY ...
      POOL_dollars_EA_sum_CCY ...
      POOL_dollars_HC_sum_CCY ...
      POOL_percent_BA_sum_CCY ...
      POOL_percent_AA_sum_CCY ...
      POOL_percent_EA_sum_CCY ...
      POOL_percent_HC_sum_CCY ...
      POOL_dollars_sum_CCY ...
      POOL_dollars_sum_CCY_sort ...
      POOL_percent_sum_CCY ...
      POOL_percent_sum_CCY_sort ...
      POOL_local_BA_sum_CCY ...
      POOL_local_AA_sum_CCY ...
      POOL_local_EA_sum_CCY ...
      POOL_local_HC_sum_CCY ...
      POOL_local_sum_CCY ...
      POOL_local_sum_CCY_sort ...
      I_AA ...
      I_BA
  
      

%% POOL
%  BALANCE SHEET by AssetType
import mlreportgen.dom.*;
filename=strcat('POOL_BalanceSheet_AssetType_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_BalanceSheet_AssetType.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

% BA
[POOL_types_BA_unique,~,~] = unique(POOL_types_BA);
 POOL_types_BA_unique      =        POOL_types_BA_unique';
                TP_BA      = length(POOL_types_BA_unique);

for i=1:TP_BA
    STRFIND_BA=strfind(POOL_types_BA,POOL_types_BA_unique{i});
    I_BA=find(~cellfun(@isempty,STRFIND_BA));

    POOL_dollars_BA_sum_AssetType(i,1)=sum(POOL_dollars_BA(END,end,I_BA),3);
    POOL_percent_BA_sum_AssetType(i,1)=sum(POOL_percent_BA(END,end,I_BA),3);
end

% AA
[POOL_types_AA_unique,~,~] = unique(POOL_types_AA);
 POOL_types_AA_unique      =        POOL_types_AA_unique';
 TP_AA  = length(POOL_types_AA_unique);

for i=1:TP_AA
    STRFIND_AA=strfind(POOL_types_AA,POOL_types_AA_unique{i});
    I_AA=find(~cellfun(@isempty,STRFIND_AA));

    POOL_dollars_AA_sum_AssetType(i,1)=sum(POOL_dollars_AA(END,end,I_AA),3);
    POOL_percent_AA_sum_AssetType(i,1)=sum(POOL_percent_AA(END,end,I_AA),3);
end

% EA
POOL_types_EA_unique = {'ADMIN COSTS Prepd/Accrd'};
POOL_dollars_EA_sum_AssetType(1,1)=sum(POOL_dollars_EA(END,end,:),3);
POOL_percent_EA_sum_AssetType(1,1)=sum(POOL_percent_EA(END,end,:),3);

% HC
POOL_types_HC_unique               = {'HELD CUSTODY, SubRed'};
POOL_dollars_HC_sum_AssetType(1,1) = POOL_dollars_HC_sub(END,end) + POOL_dollars_HC_red(END,end);
POOL_percent_HC_sum_AssetType(1,1) = POOL_percent_HC_sub(END,end) + POOL_percent_HC_red(END,end);

% BA|AA|EA|HC aggregate
POOL_types_unique         =[POOL_types_BA_unique;...
                            POOL_types_AA_unique;...
                            POOL_types_EA_unique;...
                            POOL_types_HC_unique];

POOL_dollars_sum_AssetType=[POOL_dollars_BA_sum_AssetType;...
                            POOL_dollars_AA_sum_AssetType;...
                            POOL_dollars_EA_sum_AssetType;...
                            POOL_dollars_HC_sum_AssetType];

POOL_percent_sum_AssetType=[POOL_percent_BA_sum_AssetType;...
                            POOL_percent_AA_sum_AssetType;...
                            POOL_percent_EA_sum_AssetType;...
                            POOL_percent_HC_sum_AssetType];

[POOL_dollars_sum_AssetType_sort,I_sort]=sort(POOL_dollars_sum_AssetType,'descend');

          POOL_types_unique_sort=POOL_types_unique(I_sort);
 POOL_percent_sum_AssetType_sort=POOL_percent_sum_AssetType(I_sort) ;
                             TP =length(POOL_types_unique_sort);
 
for i=1:TP
    CONTENT{i,2}=numberFormatter(POOL_dollars_sum_AssetType_sort(i),'$###,###');
    CONTENT{i,1}=numberFormatter(POOL_percent_sum_AssetType_sort(i),'##.#%');
          RowName{i}  =POOL_types_unique_sort{i};
end
    CONTENT{TP+1,2}=numberFormatter(sum(POOL_dollars_sum_AssetType_sort),'$###,###');
    CONTENT{TP+1,1}=numberFormatter(sum(POOL_percent_sum_AssetType_sort),'##.#%');
          RowName{TP+1}='Total.............................................';

TabObj = cell2table(CONTENT,'RowNames',RowName);
TabObj.Properties.VariableNames = {'percent' 'USDval'};

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE';        value=TabObj;                                       append(DocObj, TabObj);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj CONTENT TabObj RowName PDFfile DOCfile
clear POOL_dollars_BA_sum_AssetType ...
      POOL_dollars_AA_sum_AssetType ...
      POOL_dollars_EA_sum_AssetType ...
      POOL_dollars_HC_sum_AssetType ...
      POOL_percent_BA_sum_AssetType ...
      POOL_percent_AA_sum_AssetType ...
      POOL_percent_EA_sum_AssetType ...
      POOL_percent_HC_sum_AssetType ...
      POOL_dollars_sum_AssetType ...
      POOL_dollars_sum_AssetType_sort ...
      POOL_percent_sum_AssetType ...
      POOL_percent_sum_AssetType_sort ...
      POOL_types_unique ...
      POOL_types_unique_sort ...
      POOL_types_BA_unique ...
      POOL_types_AA_unique ...
      POOL_types_EA_unique ...
	  POOL_types_HC_unique ...
      TP TP_AA TP_BA
      
%% POOL
%  BALANCE SHEET by CounterParty
import mlreportgen.dom.*;
filename=strcat('POOL_BalanceSheet_CounterParty_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_BalanceSheet_CounterParty.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

% BA|AA|EA|HC stack
  POOL_names_BAAAEAHC=[POOL_names_BA,POOL_names_AA,'Prepd/Accrd Admin Costs','HELD CUSTODY']';
POOL_dollars_BAAAEAHC=[reshape(POOL_dollars_BA(END,end,:),BA,1);reshape(POOL_dollars_AA(END,end,:),AA,1);sum(POOL_dollars_EA(END,end,:),3);POOL_dollars_HC_sub(END,end)+POOL_dollars_HC_red(END,end)];
POOL_percent_BAAAEAHC=[reshape(POOL_percent_BA(END,end,:),BA,1);reshape(POOL_percent_AA(END,end,:),AA,1);sum(POOL_percent_EA(END,end,:),3);POOL_percent_HC_sub(END,end)+POOL_percent_HC_red(END,end)];

% Uniqe CounterParties
[POOL_names_BAAAEAHC_unique,~,~] = unique(POOL_names_BAAAEAHC);
 POOL_names_BAAAEAHC_unique      =        POOL_names_BAAAEAHC_unique';
                              CP = length(POOL_names_BAAAEAHC_unique);

for i=1:CP
    STRFIND=strfind(POOL_names_BAAAEAHC,POOL_names_BAAAEAHC_unique{i});
    I=find(~cellfun(@isempty,STRFIND));
    POOL_dollars_sum_CounterParty(i,1)=sum(POOL_dollars_BAAAEAHC(I));
    POOL_percent_sum_CounterParty(i,1)=sum(POOL_percent_BAAAEAHC(I));
end

[POOL_dollars_sum_CounterParty_sort,I_sort] = sort(POOL_dollars_sum_CounterParty,'descend');
 POOL_names_BAAAEAHC_unique_sort            = POOL_names_BAAAEAHC_unique(   I_sort);
 POOL_percent_sum_CounterParty_sort         = POOL_percent_sum_CounterParty(I_sort) ;
 
for i=1:CP
    CONTENT{i,2}=numberFormatter(POOL_dollars_sum_CounterParty_sort(i),'$###,###');
    CONTENT{i,1}=numberFormatter(POOL_percent_sum_CounterParty_sort(i),'##.#%');
          RowName{i}  =POOL_names_BAAAEAHC_unique_sort{i};
end
    CONTENT{CP+1,2}=numberFormatter(sum(POOL_dollars_sum_CounterParty_sort),'$###,###');
    CONTENT{CP+1,1}=numberFormatter(sum(POOL_percent_sum_CounterParty_sort),'##.#%');
          RowName{CP+1}='Total.............................................';
TabObj = cell2table(CONTENT,'RowNames',RowName);
TabObj.Properties.VariableNames = {'percent' 'USDval'};

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE';        value=TabObj;                                       append(DocObj, TabObj);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj CONTENT TabObj RowName PDFfile DOCfile
clear POOL_names_BAAAEAHC_unique ...
      POOL_names_BAAAEAHC_unique_sort ...
      POOL_dollars_sum_CounterParty ...
      POOL_dollars_sum_CounterParty_sort ...
      POOL_percent_sum_CounterParty ...
      POOL_percent_sum_CounterParty_sort ...
      CP

%% POOL
%  BALANCE SHEET by TITLE
import mlreportgen.dom.*;
filename=strcat('POOL_BalanceSheet_TITLE_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_BalanceSheet_TITLE.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

% BA|AA|EA|HC stack
  POOL_title_BAAAEAHC=[POOL_title_BA,POOL_title_AA,POOL_title_EA,'RapFlagBVI','RapFlagBVI']';
  POOL_names_BAAAEAHC=[POOL_names_BA,POOL_names_AA,POOL_names_EA,'HeldCustodySUBSCRIPTIONS','HeldCustodyREDEMPTION']';
    POOL_ccy_BAAAEAHC=[POOL_ccy_BA  ,POOL_ccy_AA  ,POOL_ccy_EA  ,'USD','USD']';
POOL_dollars_BAAAEAHC=[reshape(POOL_dollars_BA(END,end,:),BA,1);reshape(POOL_dollars_AA(END,end,:),AA,1);reshape(POOL_dollars_EA(END,end,:),EA,1);POOL_dollars_HC_sub(END,end);POOL_dollars_HC_red(END,end)];
POOL_percent_BAAAEAHC=[reshape(POOL_percent_BA(END,end,:),BA,1);reshape(POOL_percent_AA(END,end,:),AA,1);reshape(POOL_percent_EA(END,end,:),EA,1);POOL_percent_HC_sub(END,end);POOL_percent_HC_red(END,end)];

% Uniqe CounterParties
[POOL_title_BAAAEAHC_unique,~,~] = unique(POOL_title_BAAAEAHC);
 POOL_title_BAAAEAHC_unique      =        POOL_title_BAAAEAHC_unique';
                      Ttl      = length(POOL_title_BAAAEAHC_unique);

for i=1:Ttl
    STRFIND=strfind(POOL_title_BAAAEAHC,POOL_title_BAAAEAHC_unique{i});
    I=find(~cellfun(@isempty,STRFIND));
    POOL_dollars_BAAAEAHC_TITLE{i} =POOL_dollars_BAAAEAHC(I);
    POOL_percent_BAAAEAHC_TITLE{i} =POOL_percent_BAAAEAHC(I)/sum(POOL_percent_BAAAEAHC(I)); % Percent (%) of BA|AA|EA in ENTITY's equity
      POOL_names_BAAAEAHC_TITLE{i} =  POOL_names_BAAAEAHC(I);
        POOL_ccy_BAAAEAHC_TITLE{i} =    POOL_ccy_BAAAEAHC(I);
    
   [POOL_dollars_BAAAEAHC_TITLE{i},II]=sort(POOL_dollars_BAAAEAHC_TITLE{i},'descend');
    POOL_percent_BAAAEAHC_TITLE{i}    =POOL_percent_BAAAEAHC_TITLE{i}(II);
      POOL_names_BAAAEAHC_TITLE{i}    =  POOL_names_BAAAEAHC_TITLE{i}(II);
        POOL_ccy_BAAAEAHC_TITLE{i}    =    POOL_ccy_BAAAEAHC_TITLE{i}(II);
    
    POOL_dollars_sum_TITLE(i,1)=sum(POOL_dollars_BAAAEAHC(I));
    POOL_percent_sum_TITLE(i,1)=sum(POOL_percent_BAAAEAHC(I)); % Percent (%) of ENTITY in POOL equity
end
clear I II STRFIND

for i=1:Ttl
    CONTENT{i,2}=numberFormatter(POOL_dollars_sum_TITLE(i),'$###,###');
    CONTENT{i,1}=numberFormatter(POOL_percent_sum_TITLE(i),'##.#%');
    RowName{i}  =POOL_title_BAAAEAHC_unique{i};
end
    CONTENT{Ttl+1,2}=numberFormatter(sum(POOL_dollars_sum_TITLE),'$###,###');
    CONTENT{Ttl+1,1}=numberFormatter(sum(POOL_percent_sum_TITLE),'##.#%');
    RowName{Ttl+1}='Total.............................................';
TABLE_summary = cell2table(CONTENT,'RowNames',RowName);
TABLE_summary.Properties.VariableNames = {'percent' 'USDval'};

clear CONTENT RowName

for i=1:Ttl
    II=length(POOL_dollars_BAAAEAHC_TITLE{i});
    for ii=1:II
        CONTENT{i}{ii,4}=numberFormatter(POOL_dollars_BAAAEAHC_TITLE{i}(ii),'$###,###');
        CONTENT{i}{ii,3}=numberFormatter(POOL_percent_BAAAEAHC_TITLE{i}(ii),'##.#%');
        CONTENT{i}{ii,2}  =  POOL_ccy_BAAAEAHC_TITLE{i}{ii};
        CONTENT{i}{ii,1}  =POOL_names_BAAAEAHC_TITLE{i}{ii};
    end
        CONTENT{i}{II+1,4}=numberFormatter(sum(POOL_dollars_BAAAEAHC_TITLE{i}),'$###,###');
        CONTENT{i}{II+1,3}=numberFormatter(sum(POOL_percent_BAAAEAHC_TITLE{i}),'##.#%');
        CONTENT{i}{II+1,2}  ='---';
        CONTENT{i}{II+1,1}  ='Total.............................................';
        
        TABLE_name=strcat('TABLE_',POOL_title_BAAAEAHC_unique{i});
        TABLE     =cell2table(CONTENT{i});
        VarNames  ={'Name','LocalCCY','Percent','DollarValue'};
        eval(strcat(TABLE_name,'=TABLE;'));
        
        eval(strcat(TABLE_name,'.Properties.VariableNames = VarNames;'));
clear CONTENT RowName TABLE_name TABLE_name VarNames
end

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';             value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE_summary';    value=TABLE_summary;                                append(DocObj,value);
              case 'TABLE_AviRap';     value=TABLE_AviRap;                                 append(DocObj,value);
              case 'TABLE_RapFlagBVI'; value=TABLE_RapFlagBVI;                             append(DocObj,value);
              case 'TABLE_RapFlagLLC'; value=TABLE_RapFlagLLC;                             append(DocObj,value);
              case 'TABLE_RapFlagLLP'; value=TABLE_RapFlagLLP;                             append(DocObj,value);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj PDFfile DOCfile
clear TABLE TABLE_name TABLE_summary TABLE_AviRap TABLE_RapFlagBVI TABLE_RapFlagLLC TABLE_RapFlagLLP
clear POOL_dollars_BAAAEAHC ...
      POOL_percent_BAAAEAHC ...
      POOL_title_BAAAEAHC_unique ...
      POOL_dollars_BAAAEAHC_TITLE ...
      POOL_percent_BAAAEAHC_TITLE ...
      POOL_names_BAAAEAHC_TITLE ...
      POOL_ccy_BAAAEAHC_TITLE ...
      POOL_dollars_sum_TITLE ...
      POOL_percent_sum_TITLE

%% POOL
%  INCOME STATEMENT by BA,AA,EA
import mlreportgen.dom.*;
filename=strcat('POOL_IncomeStatement_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_IncomeStatement.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

% AGGREGATES
POOL_dollars_INCOME =sum(sum(POOL_dollars(BEG:END,[03 05 10 14 16 18]),1),2);   % INCOME (BA+AA+EA)            Taxable Income/Expense
POOL_dollars_CPGAIN =sum(sum(POOL_dollars(BEG:END,[06 07 08 09 11 12 13 15 17 19]),1),2);   % CAP.GN (BA+AA+EA+Depo+With)  FX translation, Cap.Gn residuals, Cap.Gn (AA)
POOL_dollars_TTGAIN =POOL_dollars_INCOME + POOL_dollars_CPGAIN;                       % TOTAL GAIN

% POOL INCOME
CONTENT{1,2}=numberFormatter(POOL_dollars_INCOME                    ,'$###,###'); % DOLLARS
CONTENT{1,1}=numberFormatter(POOL_dollars_INCOME/POOL_dollars_TTGAIN,'###.##%');  % PERCENT
RowName{1}='INCOME(taxable)';
TabObj_INCOME = cell2table(CONTENT,'RowNames',RowName);
TabObj_INCOME.Properties.VariableNames = {'PercentTOTGAIN' 'INCOME'};
clear CONTENT RowName

% POOL CAPITAL GAIN (TOTAL)
CONTENT{1,2}=numberFormatter(POOL_dollars_CPGAIN                    ,'$###,###'); % DOLLARS
CONTENT{1,1}=numberFormatter(POOL_dollars_CPGAIN/POOL_dollars_TTGAIN,'###.##%');  % PERCENT
RowName{1}='CapitalGain';
TabObj_CPGAIN = cell2table(CONTENT,'RowNames',RowName);
TabObj_CPGAIN.Properties.VariableNames = {'PercentTOTGAIN' 'CapitalGain'};
clear CONTENT RowName

% POOL TOTAL GAIN
CONTENT{1,2}=numberFormatter( POOL_dollars_TTGAIN                                             ,'$###,###'); % DOLLARS
CONTENT{1,1}=numberFormatter((POOL_dollars_INCOME+POOL_dollars_CPGAIN)/POOL_dollars_TTGAIN,'###.##%');  % PERCENT
RowName{1}='TOTAL GAIN';
TabObj_TTGAIN = cell2table(CONTENT,'RowNames',RowName);
TabObj_TTGAIN.Properties.VariableNames = {'PercentTOTGAIN' 'TotalGain'};
clear CONTENT RowName

% INCOME by BA
[POOL_dollars_BA_sort,I_BA_sort]=sort(reshape(sum(POOL_dollars_BA(BEG:END,2,:),1),[BA,1]),'descend');  % DOLLARS
 POOL_percent_BA_sort = POOL_dollars_BA_sort / sum(POOL_dollars_BA_sort);                              % PERCENT
 
  POOL_ccy_BA_sort =   POOL_ccy_BA(I_BA_sort);
POOL_names_BA_sort = POOL_names_BA(I_BA_sort);
POOL_types_BA_sort = POOL_types_BA(I_BA_sort);
  POOL_IDs_BA_sort =   POOL_IDs_BA(I_BA_sort);
POOL_title_BA_sort = POOL_title_BA(I_BA_sort);

for ba=1:BA
    CONTENT{ba,6}=numberFormatter(POOL_dollars_BA_sort(ba),'$###,###');    % DOLLARS
    CONTENT{ba,5}=numberFormatter(POOL_percent_BA_sort(ba),'##.#%');       % PERCENT
    CONTENT{ba,4}=POOL_title_BA_sort{ba};
    CONTENT{ba,3}=  POOL_IDs_BA_sort{ba};
    CONTENT{ba,2}=  POOL_ccy_BA_sort{ba};
    CONTENT{ba,1}=POOL_types_BA_sort{ba};
    RowName{ba}=sprintf('(BA%02d) %s',ba,POOL_names_BA_sort{ba});
end
    CONTENT{BA+1,6}=numberFormatter(sum(POOL_dollars_BA_sort)                    ,'$###,###'); % DOLLARS
    CONTENT{BA+1,5}=numberFormatter(sum(POOL_dollars_BA_sort)/POOL_dollars_TTGAIN,'##.#%');    % PERCENT
    CONTENT{BA+1,4}='---';
    CONTENT{BA+1,3}='---';
    CONTENT{BA+1,2}='---';
    CONTENT{BA+1,1}='---';
    RowName{BA+1}='(BA99) Total BA';
TabObj_BA = cell2table(CONTENT,'RowNames',RowName);
TabObj_BA.Properties.VariableNames = {'AssetType' 'CCY' 'AccountID' 'AccountHolder' 'PercentBA' 'BA_INCOME'};
clear CONTENT RowName

% INCOME by AA
% CAP.GN by AA
FNDinc = strfind(POOL_CGIswitch_AA,'INCOME'); for aa=1:AA; if ~isempty(FNDinc{aa}); INDinc(aa)=aa; end; end; INDinc(find(INDinc==0))=[]; AAinc=length(INDinc);
FNDcpg = strfind(POOL_CGIswitch_AA,'CAP.GN'); for aa=1:AA; if ~isempty(FNDcpg{aa}); INDcpg(aa)=aa; end; end; INDcpg(find(INDcpg==0))=[]; AAcpg=length(INDcpg);

% CHOOSE income/cap.gn AAs
POOL_dollars_AAinc = POOL_dollars_AA(BEG:END,[2 7],INDinc); POOL_ccy_AAinc = POOL_ccy_AA(INDinc); POOL_names_AAinc = POOL_names_AA(INDinc); POOL_types_AAinc = POOL_types_AA(INDinc); POOL_IDs_AAinc = POOL_IDs_AA(INDinc); POOL_title_AAinc = POOL_title_AA(INDinc);
POOL_dollars_AAcpg = POOL_dollars_AA(BEG:END,[3 8],INDcpg); POOL_ccy_AAcpg = POOL_ccy_AA(INDcpg); POOL_names_AAcpg = POOL_names_AA(INDcpg); POOL_types_AAcpg = POOL_types_AA(INDcpg); POOL_IDs_AAcpg = POOL_IDs_AA(INDcpg); POOL_title_AAcpg = POOL_title_AA(INDcpg);

% SORT within income/cap.gn AAs
[POOL_dollars_AAinc_sort,I_AAinc_sort]=sort(reshape(sum(sum(POOL_dollars_AAinc,1),2),[AAinc,1]),'descend'); POOL_percent_AAinc_sort = POOL_dollars_AAinc_sort / sum(POOL_dollars_AAinc_sort);
[POOL_dollars_AAcpg_sort,I_AAcpg_sort]=sort(reshape(sum(sum(POOL_dollars_AAcpg,1),2),[AAcpg,1]),'descend'); POOL_percent_AAcpg_sort = POOL_dollars_AAcpg_sort / sum(POOL_dollars_AAcpg_sort);

POOL_ccy_AAinc_sort = POOL_ccy_AAinc(I_AAinc_sort); POOL_names_AAinc_sort = POOL_names_AAinc(I_AAinc_sort); POOL_types_AAinc_sort = POOL_types_AAinc(I_AAinc_sort); POOL_IDs_AAinc_sort = POOL_IDs_AAinc(I_AAinc_sort); POOL_title_AAinc_sort = POOL_title_AAinc(I_AAinc_sort);
POOL_ccy_AAcpg_sort = POOL_ccy_AAcpg(I_AAcpg_sort); POOL_names_AAcpg_sort = POOL_names_AAcpg(I_AAcpg_sort); POOL_types_AAcpg_sort = POOL_types_AAcpg(I_AAcpg_sort); POOL_IDs_AAcpg_sort = POOL_IDs_AAcpg(I_AAcpg_sort); POOL_title_AAcpg_sort = POOL_title_AAcpg(I_AAcpg_sort);

% TABLE for AAinc
for aa=1:AAinc
    CONTENT_inc{aa,6} =numberFormatter(POOL_dollars_AAinc_sort(aa),'$###,###');
    CONTENT_inc{aa,5} =numberFormatter(POOL_percent_AAinc_sort(aa),'##.#%');
    CONTENT_inc{aa,4} = POOL_title_AAinc_sort{aa};
    CONTENT_inc{aa,3} =   POOL_IDs_AAinc_sort{aa};
    CONTENT_inc{aa,2} =   POOL_ccy_AAinc_sort{aa};
    CONTENT_inc{aa,1} = POOL_types_AAinc_sort{aa};
    RowName_inc{aa}=sprintf('(AA%02d) %s',aa,POOL_names_AAinc_sort{aa});
end
    CONTENT_inc{AAinc+1,6}=numberFormatter(sum(POOL_dollars_AAinc_sort)                    ,'$###,###');  % DOLLARS
    CONTENT_inc{AAinc+1,5}=numberFormatter(sum(POOL_dollars_AAinc_sort/POOL_dollars_TTGAIN),'##.#%');     % PERCENT
    CONTENT_inc{AAinc+1,4}='---';
    CONTENT_inc{AAinc+1,3}='---';
    CONTENT_inc{AAinc+1,2}='---';
    CONTENT_inc{AAinc+1,1}='---';
    RowName_inc{AAinc+1}='(AA99) Total AAinc';

    TabObj_AAinc = cell2table(CONTENT_inc,'RowNames',RowName_inc);
    TabObj_AAinc.Properties.VariableNames = {'AssetType' 'CCY' 'AccountID' 'AccountHolder' 'PercentAA' 'AA_INCOME'};
    clear CONTENT_inc RowName_inc aa
    
% TABLE for AAcpg
for aa=1:AAcpg
    CONTENT_cpg{aa,6} =numberFormatter(POOL_dollars_AAcpg_sort(aa),'$###,###');
    CONTENT_cpg{aa,5} =numberFormatter(POOL_percent_AAcpg_sort(aa),'##.#%');
    CONTENT_cpg{aa,4} = POOL_title_AAcpg_sort{aa};
    CONTENT_cpg{aa,3} =   POOL_IDs_AAcpg_sort{aa};
    CONTENT_cpg{aa,2} =   POOL_ccy_AAcpg_sort{aa};
    CONTENT_cpg{aa,1} = POOL_types_AAcpg_sort{aa};
    RowName_cpg{aa}=sprintf('(AA%02d) %s',aa,POOL_names_AAcpg_sort{aa});
end
    CONTENT_cpg{AAcpg+1,6}=numberFormatter(sum(POOL_dollars_AAcpg_sort)                        ,'$###,###');  % DOLLARS
    CONTENT_cpg{AAcpg+1,5}=numberFormatter(sum(POOL_dollars_AAcpg_sort/POOL_dollars_TTGAIN),'##.#%');     % PERCENT
    CONTENT_cpg{AAcpg+1,4}='---';
    CONTENT_cpg{AAcpg+1,3}='---';
    CONTENT_cpg{AAcpg+1,2}='---';
    CONTENT_cpg{AAcpg+1,1}='---';
    RowName_cpg{AAcpg+1}='(AA99) Total AAcpg';

    TabObj_AAcpg = cell2table(CONTENT_cpg,'RowNames',RowName_cpg);
    TabObj_AAcpg.Properties.VariableNames = {'AssetType' 'CCY' 'AccountID' 'AccountHolder' 'PercentAA' 'AA_CPGAIN'};
    clear CONTENT_cpg RowName_cpg aa

% INCOME by EA
[POOL_dollars_EA_sort,I_EA_sort]=sort(reshape(sum(POOL_dollars_EA(BEG:END,2,:),1),[EA,1]),'descend'); % DOLLARS
 POOL_percent_EA_sort = POOL_dollars_EA_sort/sum(POOL_dollars_EA_sort);                         % PERCENT
 POOL_names_EA_sort   = POOL_names_EA(I_EA_sort);
 POOL_types_EA_sort   = POOL_types_EA(I_EA_sort);
 POOL_title_EA_sort   = POOL_title_EA(I_EA_sort);
 
for ea=1:EA
    CONTENT{ea,4}=numberFormatter(POOL_dollars_EA_sort(ea),'$###,###'); % DOLLARS
    CONTENT{ea,3}=numberFormatter(POOL_percent_EA_sort(ea),'##.#%');    % PERCENT
    CONTENT{ea,2}=                POOL_title_EA_sort{ea};               % TITLE
    CONTENT{ea,1}=                POOL_names_EA_sort{ea};               % NAME
    RowName{ea}=sprintf('(EA%02d) %s',ea,POOL_types_EA_sort{ea});       % TYPE
end
    CONTENT{EA+1,4}=numberFormatter(sum(POOL_dollars_EA_sort)                    ,'$###,###'); % DOLLARS
    CONTENT{EA+1,3}=numberFormatter(sum(POOL_dollars_EA_sort)/POOL_dollars_TTGAIN,'##.#%');    % PERCENT
    CONTENT{EA+1,2}='---';
    CONTENT{EA+1,1}='---';
    RowName{EA+1}='(EA99) Total EA';
TabObj_EA = cell2table(CONTENT,'RowNames',RowName);
TabObj_EA.Properties.VariableNames = {'ServiceProvider' 'AccountHolder' 'PercentEA' 'EA_INCOME'};
clear CONTENT RowName

% SubCost and RedCost
    RowName{1}  ='SR01';                                                    % #
    CONTENT{1,1}='Subscription Costs';                                      % NAME
    CONTENT{1,2}=numberFormatter(sum(POOL_SubRed(BEG:END,2),1),'$###,###'); % DOLLARS
    CONTENT{1,3}=numberFormatter(sum(POOL_SubRed(BEG:END,2),1)/sum(sum(POOL_SubRed(BEG:END,[2 3]),1),2),'##.#%');    % PERCENT
    
    RowName{2}  ='SR02';                                                    % #
    CONTENT{2,1}='Redemption Costs';                                        % NAME
    CONTENT{2,2}=numberFormatter(sum(POOL_SubRed(BEG:END,3),1),'$###,###'); % DOLLARS
    CONTENT{2,3}=numberFormatter(sum(POOL_SubRed(BEG:END,3),1)/sum(sum(POOL_SubRed(BEG:END,[2 3]),1),2),'##.#%');    % PERCENT
    
    
    RowName{3}  ='SR99';
    CONTENT{3,1}='Total SubRed Costs';
    CONTENT{3,2}=numberFormatter(sum(sum(POOL_SubRed(BEG:END,[2 3]),1),2)                    ,'$###,###'); % DOLLARS
    CONTENT{3,3}=numberFormatter(sum(sum(POOL_SubRed(BEG:END,[2 3]),1),2)/POOL_dollars_TTGAIN,'##.#%');    % PERCENT
    
TabObj_SR = cell2table(CONTENT,'RowNames',RowName);
TabObj_SR.Properties.VariableNames = {'SubRed' 'Cost' 'PercentTOT'};
clear CONTENT RowName

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'BEG';          value=cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'END';          value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE_BA';     append(DocObj,TabObj_BA);
              case 'TABLE_AAinc';  append(DocObj,TabObj_AAinc);
              case 'TABLE_AAcpg';  append(DocObj,TabObj_AAcpg);
              case 'TABLE_EA';     append(DocObj,TabObj_EA);
              case 'TABLE_SR';     append(DocObj,TabObj_SR);    
              case 'POOL_INCOME';  append(DocObj,TabObj_INCOME);
              case 'POOL_CPGAIN';  append(DocObj,TabObj_CPGAIN);
              case 'POOL_TTGAIN';  append(DocObj,TabObj_TTGAIN);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);

clear filename template output DocObj CONTENT EndValueUSD RowName PDFfile DOCfile
clear TabObj_BA ...
      TabObj_AAinc ...
      TabObj_AAcpg ...
      TabObj_EA ...
      TabObj_SR ...
      TabObj_INCOME ...
      TabObj_CPGAIN ...
      TabObj_TTGAIN

%% POOL
%  INCOME STATEMENT by Title
import mlreportgen.dom.*;
filename=strcat('POOL_IncomeStatement_Title_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','POOL_IncomeStatement_Title.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);

% BA|AA|EA|HC stack
  POOL_title_BAAAEA = [POOL_title_BA, POOL_title_AAinc, POOL_title_AAcpg, POOL_title_EA]';
  POOL_names_BAAAEA = [POOL_names_BA, POOL_names_AAinc, POOL_names_AAcpg, POOL_names_EA]';
  POOL_types_BAAAEA = [POOL_types_BA, POOL_types_AAinc, POOL_types_AAcpg, POOL_types_EA]';
    POOL_ccy_BAAAEA = [POOL_ccy_BA,   POOL_ccy_AAinc,   POOL_ccy_AAcpg,   POOL_ccy_EA  ]';
    POOL_CGI_BAAAEA = [repmat({'INCOME'},BA   ,1); ...
                       repmat({'INCOME'},AAinc,1); ...
                       repmat({'CAP.GN'},AAcpg,1); ...
                       repmat({'INCOME'},EA   ,1)];
POOL_dollars_BAAAEA =[reshape(sum(    POOL_dollars_BA(BEG:END,2,:)          ,1),BA   ,1);...
                      reshape(sum(sum(POOL_dollars_AAinc(:,:,:),2),1),AAinc,1);...
                      reshape(sum(sum(POOL_dollars_AAcpg(:,:,:),2),1),AAcpg,1);...
                      reshape(sum(    POOL_dollars_EA(BEG:END,2,:)          ,1),EA   ,1)];

% Unique Title Holder
[POOL_title_BAAAEA_unique,~,~] = unique(POOL_title_BAAAEA);
 POOL_title_BAAAEA_unique      =        POOL_title_BAAAEA_unique';
                Ttl            = length(POOL_title_BAAAEA_unique);

for i=1:Ttl
    STRFIND=strfind(POOL_title_BAAAEA,POOL_title_BAAAEA_unique{i});
    I=find(~cellfun(@isempty,STRFIND));
    dollars_Title=POOL_dollars_BAAAEA(I);
    percent_Title=POOL_dollars_BAAAEA(I)/sum(POOL_dollars_BAAAEA(I));
      title_Title=  POOL_title_BAAAEA(I);
      names_Title=  POOL_names_BAAAEA(I);
      types_Title=  POOL_types_BAAAEA(I);
        ccy_Title=    POOL_ccy_BAAAEA(I);
        CGI_Title=    POOL_CGI_BAAAEA(I);
   [dollars_Title,I_sort]=sort(dollars_Title,'descend');
    percent_Title=percent_Title(I_sort);
      title_Title=  title_Title(I_sort);
      names_Title=  names_Title(I_sort);
      types_Title=  types_Title(I_sort);
        ccy_Title=    ccy_Title(I_sort);
        CGI_Title=    CGI_Title(I_sort);
        for ii=1:length(names_Title)
        CONTENT{ii,7}=numberFormatter(dollars_Title(ii),'$###,###'); % DOLLARS
        CONTENT{ii,6}=numberFormatter(percent_Title(ii),'##.#%');    % PERCENT
        CONTENT{ii,5}=                    CGI_Title{ii};             % CAP.GN / INCOME (CGI)
        CONTENT{ii,4}=                  types_Title{ii};             % Types
        CONTENT{ii,3}=                    ccy_Title{ii};             % CCY
        CONTENT{ii,2}=                  names_Title{ii};             % Names
        CONTENT{ii,1}=                  title_Title{1};              % Title
        end
        CONTENT{end+1,7}=numberFormatter(sum(dollars_Title)                       ,'$###,###'); % DOLLARS
        CONTENT{end  ,6}=numberFormatter(sum(dollars_Title)/sum(POOL_dollars_BAAAEA(:)),'##.#%');    % PERCENT
        CONTENT{end  ,5}=                  '---';             % CAP.GN / INCOME (CGI)
        CONTENT{end  ,4}=                  '---';             % Types
        CONTENT{end  ,3}=                  '---';             % CCY
        CONTENT{end  ,2}=                  '---';             % Names
        CONTENT{end  ,1}=                  'Total';           % Total
        TabObj{i} = cell2table(CONTENT);
        clear CONTENT
        TabObj{i}.Properties.VariableNames = {'AccountHolder' 'AccountName' 'CCY' 'AssetType' 'CapGainIncome' 'PercentINCOME' 'INCOME'};
end

for i=1:Ttl
    STRFIND=strfind(POOL_title_BAAAEA,POOL_title_BAAAEA_unique{i});
    I=find(~cellfun(@isempty,STRFIND));
    dollars_Title=POOL_dollars_BAAAEA(I);
        CONTENT{i,2}=numberFormatter(sum(dollars_Title)                           ,'$###,###');
        CONTENT{i,1}=numberFormatter(sum(dollars_Title)/sum(POOL_dollars_BAAAEA(:)),'##.#%');
          RowName{i}  =POOL_title_BAAAEA_unique{i};
end
    CONTENT{end+1,2}=numberFormatter(sum(POOL_dollars_BAAAEA)                           ,'$###,###');
    CONTENT{end  ,1}=numberFormatter(sum(POOL_dollars_BAAAEA)/sum(POOL_dollars_BAAAEA(:)),'##.#%');
          RowName{end+1  }='Total.............................................';
TabObjTOT = cell2table(CONTENT,'RowNames',RowName);
TabObjTOT.Properties.VariableNames = {'PercentINCOME' 'INCOME'};
clear CONTENT RowName

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'BEG';         value=cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'END';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'TABLE_1';     value=TabObj{1};                                    append(DocObj, value);
              case 'TABLE_2';     value=TabObj{2};                                    append(DocObj, value);
              case 'TABLE_3';     value=TabObj{3};                                    append(DocObj, value);
              case 'TABLE_4';     value=TabObj{4};                                    append(DocObj, value);    
              case 'TABLE_TOT';   value=TabObjTOT;                                    append(DocObj, value);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj CONTENT TabObj RowName PDFfile DOCfile
clear STRFIND I I_sort ...
      dollars_Title ...
      percent_Title ...
      title_Title   ...
      names_Title   ...
      types_Title   ...
      ccy_Title
      
clear POOL_title_BAAAEA  ...
      POOL_title_BAAAEA_unique ...
      POOL_names_BAAAEA  ...
      POOL_types_BAAAEA  ...
      POOL_ccy_BAAAEA    ...
      POOL_dollars_BAAAE ...
      POOL_CGI_BAAAEA ...

%% MANAGER
%  Summary Statement
import mlreportgen.dom.*;
filename=strcat('S0_SummaryStatment_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'),'_',MNGR_name);
template=strcat('reports\templates\','MANAGER_SummaryStatement.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);
moveToNextHole(DocObj);

while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              % General Information
              case 'NAME';           value = cellstr(MNGR_name);                                                                              % NAME
              case 'NUMBER';         value = cellstr('IM');                                                                                   % SERIES
              case 'IDENTIFIER';     value = cellstr(MNGR_Identifier);                                                                        % IDENTIFIER
              case 'BEG';            value = cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY'));                                                    % DATE BEG
              case 'END';            value = cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));                                                    % DATE END
              % DOLLARS
              case 'DL_BAL_beg';     value = numberFormatter(        MNGR_dollars(BEG    ,01)   ,'$###,###');                                 % BEG
              case 'DL_SUBS';        value = numberFormatter(sum(    MNGR_dollars(BEG:END,02),1),'$###,###');                                 % SUBSCRIPTIONS
              case 'DL_COSTS_subs';  value = numberFormatter(sum(    MNGR_dollars(BEG:END,03),1),'$###,###');                                 % SubCosts
              case 'DL_INCOME';      value = numberFormatter(sum(sum(MNGR_dollars(BEG:END,[05 10 14]),2),1),'$###,###');                      % INCOME
              case 'DL_CAP_GAIN';    value = numberFormatter(sum(sum(MNGR_dollars(BEG:END,[06 07 08 09 11 12 13 15 17 19]),2),1),'$###,###'); % CAP.GN
              case 'DL_COSTS_admin'; value = numberFormatter(sum(    MNGR_dollars(BEG:END,16),1),'$###,###');                                 % AdminCost
              case 'DL_COSTS_reds';  value = numberFormatter(sum(    MNGR_dollars(BEG:END,18),1),'$###,###');                                 % RedCost
              case 'DL_COSTS_setup'; value = numberFormatter(sum(    MNGR_dollars(BEG:END,21),1),'$###,###');                                 % SetupCost
              case 'DL_FEE_mgmt';    value = numberFormatter(sum(    MNGR_dollars(BEG:END,22),1),'$###,###');                                 % MgmtFee (net of finders)
              case 'DL_FEE_perf';    value = numberFormatter(sum(    MNGR_dollars(BEG:END,23),1),'$###,###');                                 % PerfFee (net of finders)
              case 'DL_REDS';        value = numberFormatter(sum(    MNGR_dollars(BEG:END,26),1),'$###,###');                                 % REDEMPTIONS
              case 'DL_BAL_end';     value = numberFormatter(        MNGR_dollars(    END,end)  ,'$###,###');                                 % END
              % SHARES
              case 'SL_BAL_beg';     value = numberFormatter(        MNGR_shares(BEG    ,01)         ,'###,###.####');                        % BEG
              case 'SL_SUBS';        value = numberFormatter(sum(    MNGR_shares(BEG:END,02),1),'###,###.####');                              % SUBSCRIPTIONS
              case 'SL_COMP_shares'; value = numberFormatter(sum(sum(MNGR_shares(BEG:END,4:6),2),1),'###,###.####');                          % Compensation Shares
              case 'SL_COSTS_setup'; value = numberFormatter(sum(    MNGR_shares(BEG:END,04),1),'###,###.####');                              % o/w   SetupCosts
              case 'SL_FEE_mgmt';    value = numberFormatter(sum(    MNGR_shares(BEG:END,05),1),'###,###.####');                              % o/w   MgmtFee (net of finders)
              case 'SL_FEE_perf';    value = numberFormatter(sum(    MNGR_shares(BEG:END,06),1),'###,###.####');                              % o/w   PerfFee (net of finders)
              case 'SL_REDS';        value = numberFormatter(sum(    MNGR_shares(BEG:END,07),1),'###,###.####');                              % REDEMPTIONS
              case 'SL_BAL_end';     value = numberFormatter(        MNGR_shares(    END,08)   ,'###,###.####');                              % END
              % PRICES
              case 'PRICE_beg';      value = numberFormatter(         MNGR_prices(BEG    ,01),'$###,###.##');                                  % BEG
              case 'PRICE_end';      value = numberFormatter(         MNGR_prices(    END,02),'$###,###.##');                                  % END
          end
          append(DocObj,value{:});
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj PDFfile DOCfile

%% MANAGER
%  HELD in CUSTODY
import mlreportgen.dom.*;
filename=strcat('S0_HeldCustody_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'),'_',MNGR_name);
template=strcat('reports\templates\','MANAGER_HeldCustody.dotx');
output=strcat('reports\',filename);
DocObj = Document(output,'docx',template);
moveToNextHole(DocObj);

while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              % General Information
              case 'NAME';           value = cellstr(MNGR_name);                                                                            % NAME
              case 'NUMBER';         value = cellstr('IM');                                                                                 % SERIES
              case 'IDENTIFIER';     value = cellstr(MNGR_Identifier);                                                                      % IDENTIFIER
              case 'BEG';            value = cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY'));                                                  % DATE BEG
              case 'END';            value = cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));                                                  % DATE END
              % DEPOSITS not yet subscribed
              case 'SUB_BEG_BALANCE';   value = numberFormatter(    -MNGR_dollars_HC_sub(BEG    ,01)   ,'$###,###');                           % BEG
              case 'SUB_DEPOSITS';      value = numberFormatter(sum(-MNGR_dollars_HC_sub(BEG:END,02),1),'$###,###');                           % DEPOSITS
              case 'SUB_SUBSCRIPTIONS'; value = numberFormatter(sum(-MNGR_dollars_HC_sub(BEG:END,03),1),'$###,###');                           % SUBSCRIPTIONS
              case 'SUB_END_BALANCE';   value = numberFormatter(    -MNGR_dollars_HC_sub(    END,04)   ,'$###,###');                           % END
              % Redemptions not yet PAID-OUT
              case 'RED_BEG_BALANCE';   value = numberFormatter(    -MNGR_dollars_HC_red(BEG    ,01)   ,'$###,###');                           % BEG
              case 'RED_REDEMPTIONS';   value = numberFormatter(sum(-MNGR_dollars_HC_red(BEG:END,02),1),'$###,###');                           % DEPOSITS
              case 'RED_PAYOUTS';       value = numberFormatter(sum(-MNGR_dollars_HC_red(BEG:END,03),1),'$###,###');                           % SUBSCRIPTIONS
              case 'RED_END_BALANCE';   value = numberFormatter(    -MNGR_dollars_HC_red(    END,04)   ,'$###,###');                           % END
          end
          append(DocObj,value{:});
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);
clear filename template output DocObj PDFfile DOCfile


%% MANAGER
%  Share Price
% filename=strcat('S0_SharePrice_01Jan2015_',datestr(DATES(END,1),'DDmmmYYYY'),'_',MANAGER_Identifier);
% file    =[str2num(datestr(DATES(:,1),'YYYYmmDD')),MANAGER_prices];
% xlswrite(strcat('\reports\',filename),file);

%% MANAGER
%  Tax Statement
if DATES_datevec(BEG,2)==1 && DATES_datevec(BEG,3)==1 && DATES_datevec(END,2)==12 && DATES_datevec(END,3)==31 && DATES_datevec(BEG,1)==DATES_datevec(END,1)
    
    import mlreportgen.dom.*;
    filename=strcat('S0_TaxStatement_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'),'_',MNGR_name);
    template=strcat('reports\templates\','MANAGER_TaxStatement_CH.dotx');
    output=strcat('reports\',filename);
    DocObj = Document(output,'docx',template);
    
    INCOME_net      =sum(sum(MNGR_dollars(BEG:END,[03 05 10 14 16 18]),2),1);   % SubCost
                                                                                % BA: income
                                                                                % AA: Accrual, WriteOff (INCOME)
                                                                                % EA: Expense (anyways zero)
                                                                                % RedCost
    INCOME_PerShare =INCOME_net / MNGR_shares(END,end);
    WEALTH_PerShare =MNGR_prices(END,end);
    
    moveToNextHole(DocObj);
    while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'NAME';              value = cellstr(MNGR_name);                                                             % NAME
              case 'NUMBER';            value = cellstr('IM');                                                                  % SERIES
              case 'IDENTIFIER';        value = cellstr(MNGR_Identifier);                                                       % IDENTIFIER 
              case 'BEG';               value = cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY'));                                   % DATE BEG
              case 'END';               value = cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));                                   % DATE END    
              case 'INCOME_net';        value = numberFormatter(INCOME_net     ,'$###,###');                                    % Taxable Income (passive)
              case 'INCOME_PerShare';   value = numberFormatter(INCOME_PerShare,'$###,###.#####');                              % Taxable Income (passive) PER SHARE	
%               case 'AdminCosts';       value=numberFormatter(sum(    MNGR_dollars(BEG:END,16),1),'$###,###');                 % AdminCosts
              case 'SetupCosts';        value=numberFormatter(sum(    MNGR_dollars(BEG:END,21),1),'$###,###');                  % SetupCosts
              case 'MgmtFee';           value=numberFormatter(sum(    MNGR_dollars(BEG:END,22),1),'$###,###');                  % MgmtFee (net of finder)
              case 'PerfFee';           value=numberFormatter(sum(    MNGR_dollars(BEG:END,23),1),'$###,###');                  % PerfFee (net of finder)
              case 'ActiveIncome';      value=numberFormatter(sum(sum(MNGR_dollars(BEG:END,[21 22 23]),2),1),'$###,###');       % Taxable Income (active)
              case 'BAL_end';           value=numberFormatter(        MNGR_dollars(END,end),'$###,###');                        % Taxable Wealth in USD
              case 'SHARES_number';     value=numberFormatter(        MNGR_shares( END,end),'####.#####');                      % Taxable Wealth in Shares
              case 'WEALTH_PerShare';   value=numberFormatter(        WEALTH_PerShare      ,'$###,###.#####');                  % Taxable Wealth Per Share
          end
          append(DocObj,value{:});
          moveToNextHole(DocObj);
    end
    close(DocObj);
    PDFfile=strcat(fullfile(pwd,'reports\',filename));
    DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
    doc2pdf(DOCfile,PDFfile);
    pause(0.1);
    eval_expression=strcat('delete reports\',filename,'.docx');
    eval(eval_expression);
    clear filename template output DocObj PDFfile DOCfile
    
    % SWISS TAX AUTHORITY (KursListe) FORMAT
    AUDIT_SwissTaxAuthority_INCOME{1,01}=MNGR_Identifier;                    % ISIN
    AUDIT_SwissTaxAuthority_INCOME{1,02}=[];                                 % Valor
    AUDIT_SwissTaxAuthority_INCOME{1,03}='Rapaport Flagship Limited';        % Name of Fund
    AUDIT_SwissTaxAuthority_INCOME{1,04}='USD';                              % Share Class Currency
    AUDIT_SwissTaxAuthority_INCOME{1,07}=datestr(DATES(END,1),'DD/mm/YYYY'); % Closing Date
    AUDIT_SwissTaxAuthority_INCOME{1,14}='USD';                              % Reporting Currency
    AUDIT_SwissTaxAuthority_INCOME{1,15}=max(INCOME_PerShare,0);             % Taxable Income Per Share (Zero if negative)
    
    AUDIT_SwissTaxAuthority_WEALTH(1,: )=AUDIT_SwissTaxAuthority_INCOME(1,:);
    AUDIT_SwissTaxAuthority_WEALTH{1,15}=WEALTH_PerShare;                    % Taxable Wealth Per Share
    
    clear INCOME_net ...
          INCOME_PerShare ...
          WEALTH_PerShare
end

%% MANAGER
%  Share Ledger Transactions
import mlreportgen.dom.*;
filename=strcat('S0_SharesTX_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'));
template=strcat('reports\templates\','MANAGER_SharesTX.dotx');
output=strcat('reports\',filename);

DocObj = Document(output,'docx',template);

IND=find(sum(MNGR_shares(BEG:END,[2 4 5 6]),2)+abs(MNGR_shares(BEG:END,7))>=10^-8);
    DATES_yyyymmmdd_Trans=DATES_yyyymmmdd(BEG:END,:); DATES_yyyymmmdd_Trans =DATES_yyyymmmdd_Trans(IND,:);
    MNGR_shares_Trans =MNGR_shares( BEG:END,:); MNGR_shares_Trans  =MNGR_shares_Trans( IND,:);
    MNGR_prices_Trans =MNGR_prices( BEG:END,:); MNGR_prices_Trans  =MNGR_prices_Trans( IND,:);
    MNGR_dollars_Trans=MNGR_dollars(BEG:END,:); MNGR_dollars_Trans =MNGR_dollars_Trans(IND,:);

TabObj = table(                     cellstr(DATES_yyyymmmdd_Trans(:,:)),...                             % DATES
                            numberFormatter(    MNGR_shares_Trans( :,01)           ,'###,###.####'),... % SHARES    BEG BALANCE
                            numberFormatter(    MNGR_shares_Trans( :,02)           ,'###,###.####'),... % SHARES    SUBSCRIPTIONS
                            numberFormatter(sum(MNGR_shares_Trans( :,[4 5 6]),2)   ,'###,###.####'),... % SHARES    CompensationShares (Mgmt+Perf+Setup)
                            numberFormatter(    MNGR_shares_Trans( :,07)           ,'###,###.####'),... % SHARES    REDEMPTIONS
                            numberFormatter(    MNGR_shares_Trans( :,08)           ,'###,###.####'),... % SHARES    END BALANCE
                            numberFormatter(    MNGR_prices_Trans( :,01)           ,'###,###.##'),...   % PRICE     BEG
                            numberFormatter(    MNGR_prices_Trans( :,02)           ,'###,###.##'),...   % PRICE     END
                            numberFormatter(    MNGR_dollars_Trans(:,01)           ,'###,###'),...      % DOLLARS   BEG BALANCE
                            numberFormatter(    MNGR_dollars_Trans(:,02)           ,'###,###'),...      % DOLLARS   SUBSCRIPTIONS
                            numberFormatter(sum(MNGR_dollars_Trans(:,[17 18 19]),2),'###,###'),...      % DOLLARS   Compensation (Mgmt+Perf+Setup)
                            numberFormatter(    MNGR_dollars_Trans(:,20)           ,'###,###'),...      % DOLLARS   REDEMPTIONS
                            numberFormatter(    MNGR_dollars_Trans(:,end)          ,'###,###'));        % DOLLARS   END BALANCE
                       
TabObj.Properties.VariableNames = {'Date' 'BegBal' 'Subs' 'Comp' 'Reds' 'EndBal' 'pBeg' 'pEnd' 'vBegBal' 'vSubs' 'vComp' 'vReds' 'vEndBal'};

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              % General Information
              case 'NAME';        value=cellstr(MNGR_name);                           append(DocObj,value{:});
              case 'NUMBER';      value=cellstr('IM');                                append(DocObj,value{:});
              case 'BEG';         value=cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              case 'END';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
              % ShareLedgerTransactions    
              case 'TABLE';       value=TabObj;                                       append(DocObj,value);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);

clear filename template output DocObj TabObj EndValueUSD RowName PDFfile DOCfile
clear MNGR_shares_Trans MNGR_prices_Trans MNGR_dollars_Trans DATES_yyyymmmdd_Trans

%% CLIENTS
%  Summary Statement
S_select=1:S;
S_select(CLTS_FullyPaidOut)=[]; % Remove Fully Paid Out Series

for s=S_select
    import mlreportgen.dom.*;
    filename=strcat('S',num2str(s),'_SummaryStatement_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'),'_',CLTS_Names(s));
    
    if strcmp(CLTS_finderYN{s},'FINDER'); template=strcat('reports\templates\','CLIENTS_SummaryStatement_FINDER.dotx'); end
    if strcmp(CLTS_finderYN{s},'NOT');    template=strcat('reports\templates\','CLIENTS_SummaryStatement.dotx');        end
    
    output=strcat('reports\',filename{:});
    DocObj = Document(output,'docx',template);
    
    moveToNextHole(DocObj);
    while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              % General Information
              case 'NAME';          value = cellstr(CLTS_Names(s));                                                                             % NAME
              case 'NUMBER';        value = cellstr(num2str(s));                                                                                % SERIES
              case 'IDENTIFIER';    value = cellstr(CLTS_Identifier(s));                                                                        % IDENTIFIER
              case 'BEG';           value = cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY'));                                                       % DATE BEG
              case 'END';           value = cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));                                                       % DATE END
              % DOLLAR Ledger
              case 'DL_BAL_beg';    value = numberFormatter(        CLTS_dollars(BEG    ,01,s)   ,'$###,###');                                  % BEG BALANCE
              case 'DL_SUBS';       value = numberFormatter(sum(    CLTS_dollars(BEG:END,02,s),1),'$###,###');                                  % SUBSCRIPTIONS
              case 'DL_SubCosts';   value = numberFormatter(sum(    CLTS_dollars(BEG:END,03,s),1),'$###,###');                                  % SubCosts
              case 'DL_INCOME';     value = numberFormatter(sum(sum(CLTS_dollars(BEG:END,[05 10 14],s),2),1),'$###,###');                    % Taxable Income (ex   AdminCosts)
              case 'DL_CAP_GAIN';   value = numberFormatter(sum(sum(CLTS_dollars(BEG:END,[06 07 08 09 11 12 13 15 17 19],s),2),1),'$###,###');  % Capital Gain   (incl Deposits, Withdrawals, Reversals, and AddBacks)
              case 'DL_AdminCosts'; value = numberFormatter(sum(    CLTS_dollars(BEG:END,16,s),1),'$###,###');                                  % AdminCosts
              case 'DL_RedCosts';   value = numberFormatter(sum(    CLTS_dollars(BEG:END,18,s),1),'$###,###');                                  % RedCosts
              case 'DL_SetupCosts'; value = numberFormatter(sum(    CLTS_dollars(BEG:END,21,s),1),'$###,###');                                  % SetupCosts
              case 'DL_MgmtFee';    value = numberFormatter(sum(    CLTS_dollars(BEG:END,22,s),1),'$###,###');                                  % MgmtFee (gross paid to MNGR+FNDR)
              case 'DL_PerfFee';    value = numberFormatter(sum(    CLTS_dollars(BEG:END,23,s),1),'$###,###');                                  % PerfFee (gross paid to MNGR+FNDR)
              case 'DL_FINDERS';    value = numberFormatter(sum(    CLTS_dollars(BEG:END,24,s),1),'$###,###');                                  % FinderFee received
              case 'DL_REDS';       value = numberFormatter(sum(    CLTS_dollars(BEG:END,26,s),1),'$###,###');                                  % REDEMPTIONS
              case 'DL_BAL_end';    value = numberFormatter(        CLTS_dollars(    END,end,s)  ,'$###,###');                                  % END BALANCE
              case 'DL_GAIN';       value = numberFormatter(sum(sum(CLTS_dollars(BEG:END,[3,5:19,21:24],s),2),1),'$###,###');                   % Gross Gain       (INCOME + COSTS + FEES + FINDER + CapGain)
              % SHARE Ledger
              case 'SL_BAL_beg';    value = numberFormatter(        CLTS_shares(BEG    ,01,s)   ,'###,###.####');                               % SHARES BEG BALANCE
              case 'SL_SUBS';       value = numberFormatter(sum(    CLTS_shares(BEG:END,02,s),1),'###,###.####');                               % SHARES SUBSCRIPTIONS
              case 'SL_REDS';       value = numberFormatter(sum(    CLTS_shares(BEG:END,04,s),1),'###,###.####');                               % SHARES REDEMPTIONS
              case 'SL_BAL_end';    value = numberFormatter(        CLTS_shares(    END,05,s)   ,'###,###.####');                               % SHARES END BALANCE
              % SHARE Price
              case 'PRICE_beg';     value = numberFormatter(        CLTS_prices(BEG    ,01,s)   ,'$###,###.##');                                % PRICE BEG
              case 'PRICE_end';     value = numberFormatter(        CLTS_prices(    END,02,s)   ,'$###,###.##');                                % PRICE END
              % Deposit, Subscribe, Redeem, Payout
              case 'DATE_dep';                                     value = cellstr(datestr(CLTS_dates_DEP_dtnm(s),'DD-mmm-YYYY'));                                                  % DATE deposit
              case 'DATE_sub';                                     value = cellstr(datestr(CLTS_dates_SUB_dtnm(s),'DD-mmm-YYYY'));                                                  % DATE subscription
              case 'DATE_red1';     if CLTS_dates_RED_dtnm(1,s)>0; value = cellstr(datestr(CLTS_dates_RED_dtnm(1,s),'DD-mmm-YYYY')); else; value = cellstr(''); end                 % DATE redemption 1
              case 'DATE_red2';     if CLTS_dates_RED_dtnm(2,s)>0; value = cellstr(datestr(CLTS_dates_RED_dtnm(2,s),'DD-mmm-YYYY')); else; value = cellstr(''); end                 % DATE redemption 2
              case 'DATE_redR';     if CLTS_dates_RED_dtnm(R,s)>0; value = cellstr(datestr(CLTS_dates_RED_dtnm(R,s),'DD-mmm-YYYY')); else; value = cellstr(''); end                 % DATE redemption R
              case 'DATE_pay1';     if CLTS_dates_PAY_dtnm(1,s)>0; value = cellstr(datestr(CLTS_dates_PAY_dtnm(1,s),'DD-mmm-YYYY')); else; value = cellstr(''); end                 % DATE payout 1
              case 'DATE_pay2';     if CLTS_dates_PAY_dtnm(2,s)>0; value = cellstr(datestr(CLTS_dates_PAY_dtnm(2,s),'DD-mmm-YYYY')); else; value = cellstr(''); end                 % DATE payout 2
              case 'DATE_payP';     if CLTS_dates_PAY_dtnm(P,s)>0; value = cellstr(datestr(CLTS_dates_PAY_dtnm(P,s),'DD-mmm-YYYY')); else; value = cellstr(''); end                 % DATE payout P
              case 'AMNT_dep';                                     value = numberFormatter(CLTS_Deposit( CLTS_dates_DEP_tinT(1,s),1,s),'$###,###.##');                               % AMNT deposit
              case 'AMNT_sub';                                     value = numberFormatter(CLTS_SubRed(  CLTS_dates_SUB_tinT(1,s),1,s),'$###,###.##');                               % AMNT subscription
              case 'AMNT_red1';     if CLTS_dates_RED_dtnm(1,s)>0; value = numberFormatter(-CLTS_SubRed( CLTS_dates_RED_tinT(1,s),4,s),'$###,###.##'); else; value = cellstr(''); end % AMNT redemption 1
              case 'AMNT_red2';     if CLTS_dates_RED_dtnm(2,s)>0; value = numberFormatter(-CLTS_SubRed( CLTS_dates_RED_tinT(2,s),4,s),'$###,###.##'); else; value = cellstr(''); end % AMNT redemption 2
              case 'AMNT_redR';     if CLTS_dates_RED_dtnm(R,s)>0; value = numberFormatter(-CLTS_SubRed( CLTS_dates_RED_tinT(R,s),4,s),'$###,###.##'); else; value = cellstr(''); end % AMNT redemption R
              case 'AMNT_pay1';     if CLTS_dates_PAY_dtnm(1,s)>0; value = numberFormatter(-CLTS_Payouts(CLTS_dates_PAY_tinT(1,s),1,s),'$###,###.##'); else; value = cellstr(''); end % AMNT payout 1
              case 'AMNT_pay2';     if CLTS_dates_PAY_dtnm(2,s)>0; value = numberFormatter(-CLTS_Payouts(CLTS_dates_PAY_tinT(2,s),1,s),'$###,###.##'); else; value = cellstr(''); end % AMNT payout 2
              case 'AMNT_payP';     if CLTS_dates_PAY_dtnm(P,s)>0; value = numberFormatter(-CLTS_SeriesClosed{8,s}                    ,'$###,###.##'); else; value = cellstr(''); end % AMNT payout P
              % DEPOSITS not yet subscribed
              case 'SUB_BEG_BALANCE';   value = numberFormatter(    -CLTS_dollars_HC_sub(BEG    ,01,s)   ,'$###,###');                         % BEG
              case 'SUB_DEPOSITS';      value = numberFormatter(sum(-CLTS_dollars_HC_sub(BEG:END,02,s),1),'$###,###');                         % DEPOSITS
              case 'SUB_SUBSCRIPTIONS'; value = numberFormatter(sum(-CLTS_dollars_HC_sub(BEG:END,03,s),1),'$###,###');                         % SUBSCRIPTIONS
              case 'SUB_END_BALANCE';   value = numberFormatter(    -CLTS_dollars_HC_sub(    END,04,s)   ,'$###,###');                         % END
              % REDEMPTIONS not yet PAID-OUT
              case 'RED_BEG_BALANCE';   value = numberFormatter(    -CLTS_dollars_HC_red(BEG    ,01,s)   ,'$###,###');                         % BEG
              case 'RED_REDEMPTIONS';   value = numberFormatter(sum(-CLTS_dollars_HC_red(BEG:END,02,s),1),'$###,###');                         % DEPOSITS
              case 'RED_PAYOUTS';       value = numberFormatter(sum(-CLTS_dollars_HC_red(BEG:END,03,s),1),'$###,###');                         % SUBSCRIPTIONS
              case 'RED_END_BALANCE';   value = numberFormatter(    -CLTS_dollars_HC_red(    END,04,s)   ,'$###,###');                         % END
          end
          append(DocObj,value{:});
          moveToNextHole(DocObj);
    end
    close(DocObj);
    PDFfile=strcat(fullfile(pwd,'reports\',filename));
    DOCfile=strcat(pwd,filesep,'reports\',filename{:},'.docx');
    doc2pdf(DOCfile,PDFfile{:});
    pause(0.1);
    eval_expression=strcat('delete reports\',filename,'.docx');
    eval(eval_expression{:});
    clear filename template output DocObj PDFfile DOCfile
end


%% CLIENTS
%  Share Prices
% for s=1:S
%     filename=strcat('S',num2str(s),'_SharePrice_','01Jan2015','_',datestr(DATES(END,1),'DDmmmYYYY_'),CLIENTS_Identifier{s});
%     file    =[str2num(datestr(DATES(:,1),'YYYYmmDD')),CLIENTS_prices(:,:,s)];
%     xlswrite(strcat('\reports\',filename),file);
% end

%% CLIENTS
%  Tax Statement Switzerland
if DATES_datevec(BEG,2)==1 && DATES_datevec(BEG,3)==1 && DATES_datevec(END,2)==12 && DATES_datevec(END,3)==31 && DATES_datevec(BEG,1)==DATES_datevec(END,1) % 01.JAN.XXXX to 31.DEC.XXXX
    for s=S_select
        if strcmp(CLTS_TaxResidence{s},'CH')             && ... Swiss Tax Resident
           length(DATES_datevec)>=CLTS_dates_SUB_tinT(s) % Exclude series which are not yet subscribed (e.g., deposit on 25 JAN 2022 to subscribe on 01 FEB 2022 and we're doing the 2021 Tax Statements on 04 FEB 2022 (after doing the JAN 2022 monthly report).
          
                import mlreportgen.dom.*;
                filename = strcat('S',num2str(s),'_TaxStatement_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'),'_',CLTS_Names(s));
                template = strcat('reports\templates\','CLIENTS_TaxStatement_CH.dotx');
                output   = strcat('reports\',filename{:});
                DocObj   = Document(output,'docx',template);
            
                % Series is OPEN
                if  isempty(CLTS_SeriesClosed{1,s})
                    if DATES_datevec(CLTS_dates_SUB_tinT(s),1)==DATES_datevec(END,1)                 % Opened in current year
                            N_mnths  = DATES_datevec(END,2)-DATES_datevec(CLTS_dates_SUB_tinT(s),2)+1;
                    else    N_mnths  = 12;
                    end

                    INCOME_gross        = sum(sum(CLTS_dollars(BEG:END,[03 05 10 14 18 24],s),2),1);    % SubCost
                                                                                                    % BA: income
                                                                                                    % AA: Accrual, WriteOff (INCOME)
                                                                                                    % RedCost
                                                                                                    % Finder's Received
                    EXPENSES_total      = sum(sum(CLTS_dollars(BEG:END,[16 21 22 23],s),2),1);          % AdminCosts, SetupCosts, MgmtFee, PerfFee
                    EXPENSES_MaxAllowed =  -0.015 * CLTS_dollars(END,end,s) * (N_mnths/12);
                    EXPENSES_allowed    = max(EXPENSES_MaxAllowed,EXPENSES_total);
                    EXPENSES_NotAllowed = min(EXPENSES_total-EXPENSES_allowed,0);
                    INCOME_net          = INCOME_gross + EXPENSES_allowed;
                    INCOME_PerShare     = INCOME_net/CLTS_shares(END,end,s);
                    WEALTH_PerShare     = CLTS_prices(END,end,s);
                end
            
                % Series is CLOSED
                if ~isempty(CLTS_SeriesClosed{1,s})
                    END_cls = CLTS_SeriesClosed{3,s};
                    RedAmt  =-CLTS_SeriesClosed{4,s};
                    RedShs  =-CLTS_shares(END_cls,4,s);
                    if DATES_datevec(CLTS_dates_SUB_tinT(s),1)==DATES_datevec(END,1)
                            N_mnths  = DATES_datevec(END_cls,2)-DATES_datevec(CLTS_dates_SUB_tinT(s),2)+1;
                    else    N_mnths  = DATES_datevec(END_cls,2);
                    end
                    INCOME_gross    = sum(sum(CLTS_dollars(BEG:END,[03 05 10 14 18 24],s),2),1); % For entire year.
                    EXPENSES_total  = sum(sum(CLTS_dollars(BEG:END,[16 21 21 23],s),2),1);       % For entire year.
                    EXPENSES_MaxAllowed = -0.015 * RedAmt * (N_mnths/12);
                    EXPENSES_allowed    = max(EXPENSES_MaxAllowed,EXPENSES_total);
                    EXPENSES_NotAllowed = min(EXPENSES_total-EXPENSES_MaxAllowed,0);
                    INCOME_net          = INCOME_gross + EXPENSES_allowed;
                    INCOME_PerShare     = INCOME_net/RedShs;
                    WEALTH_PerShare     = CLTS_prices(END_cls,end,s);
                end
            
                moveToNextHole(DocObj);
                while ~strcmp(DocObj.CurrentHoleId,'#end#')
                    switch DocObj.CurrentHoleId
                        case 'NAME';                value=cellstr(CLTS_Names(s));
                        case 'NUMBER';              value=cellstr(num2str(s));
                        case 'IDENTIFIER';          value=cellstr(CLTS_Identifier(s));
                        case 'BEG';                 value=cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY'));
                        case 'END';                 value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));
                        case 'INCOME_gross';        value=numberFormatter(INCOME_gross       ,'$###,###');              % Taxable Income less Direct Costs(gross)
                        case 'EXPENSES_total';      value=numberFormatter(EXPENSES_total     ,'$###,###');              % Expenses gross TOTAL 
                    	case 'EXPENSES_MaxAllowed'; value=numberFormatter(EXPENSES_MaxAllowed,'$###,###');              % Exepnses MAX ALLOWED
                        case 'DATE_SeriesClosed'
                                if isempty(CLTS_SeriesClosed{1,s})
                                                    value={'Open Series'};
                                else;               value=CLTS_SeriesClosed(2,s);
                                end
                        case 'EXPENSES_allowed';    value=numberFormatter(EXPENSES_allowed          ,'$###,###');       % Exepnses     ALLOWED
                        case 'EXPENSES_NotAllowed'; value=numberFormatter(EXPENSES_NotAllowed       ,'$###,###');       % Exepnses    NOT-ALLOWED                                          
                        case 'INCOME_net';          value=numberFormatter(INCOME_net                ,'$###,###');
                        case 'INCOME_PerShare';     value=numberFormatter(INCOME_PerShare           ,'$###,###.#####'); % Taxable Income Per Share
                        case 'BAL_end'
                            if isempty(CLTS_SeriesClosed{1,s})
                                                    value=numberFormatter(CLTS_dollars(END,end,s),'$###,###');       % Taxable Wealth in USD
                            else;                   value=numberFormatter(RedAmt                 ,'$###,###');       % Taxable Wealth in USD   (Redeemed Amount in final redemption)
                            end
                        case 'SHARES_number'
                            if isempty(CLTS_SeriesClosed{1,s})
                                                    value=numberFormatter(CLTS_shares( END,end,s),'####.#####');     % Taxable Wealth in Shares
                            else;                   value=numberFormatter(RedShs             ,'####.#####');            % Taxable Wealth in Shares(Redeemed # of Shares in final redemption)
                            end
                        case 'WEALTH_PerShare';     value=numberFormatter(WEALTH_PerShare           ,'$###,###.#####'); % Taxable Wealth PerShares
                    end
                    append(DocObj,value{:});
                    moveToNextHole(DocObj);
                end


                %       SWISS TAX AUTHORITY (KursListe) FORMAT
                        AUDIT_SwissTaxAuthority_INCOME{1+s,01}=[];                                              % ISIN
                        AUDIT_SwissTaxAuthority_INCOME{1+s,02}=CLTS_Identifier{s};                              % Valor
                        AUDIT_SwissTaxAuthority_INCOME{1+s,03}='Rapaport Flagship Limited';                     % Name of Fund
                
                        AUDIT_SwissTaxAuthority_INCOME{1+s,04}='USD';                                           % Share Class Currency                
                
                if      isempty(CLTS_SeriesClosed{1,s})
                        AUDIT_SwissTaxAuthority_INCOME{1+s,07}=datestr(DATES(END    ,1),'DD/mm/YYYY');          % Closing Date
                else;   AUDIT_SwissTaxAuthority_INCOME{1+s,07}=datestr(DATES(END_cls,1),'DD/mm/YYYY');          % Closing Date
                end
                
                        AUDIT_SwissTaxAuthority_INCOME{1+s,14}='USD';                                           % Accumulating Currency         
                
                if isempty(CLTS_SeriesClosed{1,s})
                        AUDIT_SwissTaxAuthority_INCOME{1+s,15}=max(INCOME_PerShare,0);                 % Taxable Income Per Share (Zero if negative)
                else;   AUDIT_SwissTaxAuthority_INCOME{1+s,15}=0;                                      % Taxable Income Per Share
                end
                
                AUDIT_SwissTaxAuthority_WEALTH(1+s, :)=AUDIT_SwissTaxAuthority_INCOME(1+s,:);     
                if isempty(CLTS_SeriesClosed{1,s})
                        AUDIT_SwissTaxAuthority_WEALTH{1+s,15}=WEALTH_PerShare;                        % Taxable Wealth Per Share
                else;   AUDIT_SwissTaxAuthority_INCOME{1+s,15}=0;                                      % Taxable Income Per Share
                end
                
                close(DocObj);
            
                PDFfile = strcat(fullfile(pwd,'reports\',filename));
                DOCfile = strcat(pwd,filesep,'reports\',filename{:},'.docx');
                      doc2pdf(DOCfile,PDFfile{:});
                      pause(0.1);
                eval_expression=strcat('delete reports\',filename,'.docx');
                eval(eval_expression{:});
            
                clear filename template output DocObj PDFfile DOCfile
                clear END_cls RedAmt RedShs N_mnths INCOME_gross EXPENSES_total EXPENSES_MaxAllowed EXPENSES_allowed EXPENSES_NotAllowed INCOME_net INCOME_PerShare
           
        end
    end

% EXPORT SWISS TAX AUTHORITY (KursListe) XLSX FORMAT
    
    % Remove empty rows
    DEL_ind=0;
    for ss=1:length(AUDIT_SwissTaxAuthority_INCOME)
    	if isempty(AUDIT_SwissTaxAuthority_INCOME{ss,1}) && isempty(AUDIT_SwissTaxAuthority_INCOME{ss,2})
           
           DEL_ind(end+0)=ss;
           DEL_ind(end+1)=0;

        end
    end
    DEL_ind(end)=[];
       
    AUDIT_SwissTaxAuthority_INCOME(DEL_ind,:)=[];
    AUDIT_SwissTaxAuthority_WEALTH(DEL_ind,:)=[];
    
    % Round to 4 decimal points
    for rr=1:length(AUDIT_SwissTaxAuthority_INCOME)
        AUDIT_SwissTaxAuthority_INCOME{rr,11}=round(AUDIT_SwissTaxAuthority_INCOME{rr,11},5);
        AUDIT_SwissTaxAuthority_WEALTH{rr,11}=round(AUDIT_SwissTaxAuthority_WEALTH{rr,11},5);
    end
    
    % Export to XLS
    xlswrite(strcat('\reports\AUDIT_SwissTaxAuthority_INCOME_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_to_',datestr(DATES(END,1),'DDmmmYYYY')),AUDIT_SwissTaxAuthority_INCOME);
    xlswrite(strcat('\reports\AUDIT_SwissTaxAuthority_WEALTH_',datestr(DATES(END,1),'DDmmmYYYY')),AUDIT_SwissTaxAuthority_WEALTH);
    
    clear DEL_ind_INCOME DEL_ind_WEALTH
end

%% CLIENTS
%  Share Ledger Transactions
% for s=1:S
% 
% import mlreportgen.dom.*;
% filename=strcat('S',num2str(s),'_SharesTX_',datestr(DATES(BEG,1),'DDmmmYYYY'),'_',datestr(DATES(END,1),'DDmmmYYYY'),'_',CLTS_Names(s));
% template=strcat('reports\templates\','CLIENTS_SharesTX.dotx');
% output=strcat('reports\',filename);
% 
% DocObj = Document(output{:},'docx',template);
% 
% IND=find(sum(CLTS_shares(BEG:END,2,s),2)+abs(CLTS_shares(BEG:END,4,s))>=10^-8);
% 
% if ~isempty(IND)
%     DATES_yyyymmmdd_Trans=DATES_yyyymmmdd(BEG:END,:); DATES_yyyymmmdd_Trans =DATES_yyyymmmdd_Trans(IND,:);
%     CLTS_shares_Trans =CLTS_shares( BEG:END,:,s); CLTS_shares_Trans  =CLTS_shares_Trans( IND,:);
%     CLTS_prices_Trans =CLTS_prices( BEG:END,:,s); CLTS_prices_Trans  =CLTS_prices_Trans( IND,:);
%     CLTS_dollars_Trans=CLTS_dollars(BEG:END,:,s); CLTS_dollars_Trans =CLTS_dollars_Trans(IND,:);
% 
% TABLE = table(cellstr(DATES_yyyymmmdd_Trans(:,:)),...
%                            numberFormatter(CLTS_shares_Trans( :,001),'###,###.####'),...  % BeginningBalance
%                            numberFormatter(CLTS_shares_Trans( :,002),'###,###.####'),...  % Subscriptions
%                            numberFormatter(CLTS_shares_Trans( :,004),'###,###.####'),...  % Redemptions
%                            numberFormatter(CLTS_shares_Trans( :,005),'###,###.####'),...  % EndingBalance
%                            numberFormatter(CLTS_prices_Trans( :,001),'###,###.##'),...    % BeginingPrice
%                            numberFormatter(CLTS_prices_Trans( :,002),'###,###.##'),...    % EndingPrice
%                            numberFormatter(CLTS_dollars_Trans(:,001),'###,###'),...       % BeginningBalance
%                            numberFormatter(CLTS_dollars_Trans(:,002),'###,###'),...       % Subscriptions
%                            numberFormatter(CLTS_dollars_Trans(:,026),'###,###'),...       % Redemptions
%                            numberFormatter(CLTS_dollars_Trans(:,end),'###,###'));         % EndingBalance
%                        
% TABLE.Properties.VariableNames = {'Date' 'BegBal' 'Subs' 'Reds' 'EndBal' 'pBeg' 'pEnd' 'vBegBal' 'vSubs' 'vReds' 'vEndBal'};
% 
% moveToNextHole(DocObj);
% while ~strcmp(DocObj.CurrentHoleId,'#end#')
%           switch DocObj.CurrentHoleId
%               % General Information
%               case 'NAME';        value=cellstr(CLTS_Names(s));                    append(DocObj,value{:});
%               case 'NUMBER';      value=cellstr(num2str(s));                          append(DocObj,value{:});
%               case 'BEG';         value=cellstr(datestr(DATES(BEG,1),'DD-mmm-YYYY')); append(DocObj,value{:});
%               case 'END';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY')); append(DocObj,value{:});
%               % ShareLedgerTransactions    
%               case 'TABLE';       value=TABLE;                                       append(DocObj,value);
%           end
%           moveToNextHole(DocObj);
% end
% close(DocObj);
% PDFfile=strcat(fullfile(pwd,'reports\',filename));
% DOCfile=strcat(pwd,filesep,'reports\',filename{:},'.docx');
% doc2pdf(DOCfile,PDFfile{:});
% pause(0.1);
% eval_expression=strcat('delete reports\',filename,'.docx');
% eval(eval_expression{:});
% 
% clear filename template output DocObj TABLE EndValueUSD RowName PDFfile DOCfile
% clear CLTS_shares_Trans CLTS_prices_Trans CLTS_dollars_Trans DATES_yyyymmmdd_Trans
% 
% end
% end


%% CLIENTS
%  Balance Sheet
% for s=1:S
%     import mlreportgen.dom.*;
%     filename=strcat('S',num2str(s),'_BalanceSheet_',datestr(DATES(END,1),'DDmmmYYYY'),'_',CLIENTS_Names(s));
%     template=strcat('reports\templates\','CLIENTS_BalanceSheet.dotx');
%     output=strcat('reports\',filename{:});
%     DocObj = Document(output,'docx',template);
% 
%  % AA
% [CLIENTS_dollars_EndVal_AA_sort,I_AA_EndVal_sort]=sort(reshape(CLIENTS_dollars_AA(END,end,s,:),[AA,1]),'descend');
%  CLIENTS_percent_EndVal_AA_sort                  = CLIENTS_dollars_EndVal_AA_sort/CLIENTS_dollars(END,end,s);
% 
%   POOL_ccy_AA_sort =   POOL_ccy_AA( I_AA_EndVal_sort);
% POOL_names_AA_sort = POOL_names_AA( I_AA_EndVal_sort);
% POOL_types_AA_sort = POOL_types_AA( I_AA_EndVal_sort);
%   POOL_IDs_AA_sort =   POOL_IDs_AA( I_AA_EndVal_sort);
% POOL_title_AA_sort = POOL_title_AA( I_AA_EndVal_sort);
% 
% for aa=1:AA
%     CONTENT{aa,5}= numberFormatter(CLIENTS_dollars_EndVal_AA_sort(aa),'$###,###');
%     CONTENT{aa,4}= numberFormatter(CLIENTS_percent_EndVal_AA_sort(aa),'##.#%');
%     CONTENT{aa,3}=   POOL_IDs_AA_sort{aa};
%     CONTENT{aa,2}=   POOL_ccy_AA_sort{aa};
%     CONTENT{aa,1}= POOL_types_AA_sort{aa};
%     RowName{aa,1}= sprintf('(AA%02d) %s',aa,POOL_names_AA_sort{aa});
% end
%     CONTENT{AA+1,5}=numberFormatter(sum(CLIENTS_dollars_EndVal_AA_sort),'$###,###');
%     CONTENT{AA+1,4}=numberFormatter(sum(CLIENTS_percent_EndVal_AA_sort),'##.#%');
%     CONTENT{AA+1,3}='---';
%     CONTENT{AA+1,2}='---';
%     CONTENT{AA+1,1}='---';
%     RowName{AA+1,1}='(AA99) Total AA';
%     
% TabObj_AA = cell2table(CONTENT,'RowNames',RowName);
% TabObj_AA.Properties.VariableNames = {'AssetType' 'CCY' 'AccountNumber' 'PercentEQ' 'EndValUSD'};
% clear CONTENT RowName
% 
%  % EA
% [CLIENTS_dollars_EndVal_EA_sort,I_EA_EndVal_sort]=sort(reshape(CLIENTS_dollars_EA(END,end,s,:),[EA,1]),'descend');
%  CLIENTS_percent_EndVal_EA_sort                  =CLIENTS_dollars_EndVal_EA_sort/CLIENTS_dollars(END,end,s);
% 
%   POOL_names_EA_sort = POOL_names_EA( I_EA_EndVal_sort);
%   POOL_types_EA_sort = POOL_types_EA( I_EA_EndVal_sort);
% 
% for ea=1:EA
%     CONTENT{ea,3}=numberFormatter(CLIENTS_dollars_EndVal_EA_sort(ea),'$###,###');
%     CONTENT{ea,2}=numberFormatter(CLIENTS_percent_EndVal_EA_sort(ea),'##.#%');
%     CONTENT{ea,1}=POOL_names_EA_sort{ea};
%     RowName{ea,1}=sprintf('(EA%02d) %s',ea,POOL_types_EA_sort{ea});
% end
%     CONTENT{EA+1,3}=numberFormatter(sum(CLIENTS_dollars_EndVal_EA_sort),'$###,###');
%     CONTENT{EA+1,2}=numberFormatter(sum(CLIENTS_percent_EndVal_EA_sort),'##.#%');
%     CONTENT{EA+1,1}='---';
%     RowName{EA+1,1}='(EA99) Total EA';
%         
% TabObj_EA = cell2table(CONTENT,'RowNames',RowName);
% TabObj_EA.Properties.VariableNames = {'ServiceProvider' 'PercentEQ' 'EndValUSD'};
% clear CONTENT RowName
% 
%  % TOTAL
% CONTENT{1}=numberFormatter(CLIENTS_dollars(END,end,s),'$###,###');
% RowName{1,1}='SERIES NAV';
% TabObj_NAV = cell2table(CONTENT,'RowNames',RowName);
% TabObj_NAV.Properties.VariableNames = {'EndValUSD'};
% clear EndValueUSD RowName
% 
%  % PLACE in HOLES
% moveToNextHole(DocObj);
% while ~strcmp(DocObj.CurrentHoleId,'#end#')
%           switch DocObj.CurrentHoleId
%               case 'DATE';         value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));         append(DocObj,value{:});
%               case 'TABLE_AA';     value=TabObj_AA;                                            append(DocObj,value);
%               case 'TABLE_EA';     value=TabObj_EA;                                            append(DocObj,value);
%               case 'POOL_NAV';     value=TabObj_NAV;                                           append(DocObj,value);
%           end
%           moveToNextHole(DocObj);
% end
% close(DocObj);
% PDFfile=strcat(fullfile(pwd,'reports\',filename));
% DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
% doc2pdf(DOCfile{:},PDFfile{:});
% pause(0.1);
% eval_expression=strcat('delete reports\',filename,'.docx');
% eval(eval_expression{:});
% 
% clear filename template output DocObj CONTENT TabObj_BA TabObj_AA TabObj_EA TabObj_NAV EndValueUSD RowName PDFfile DOCfile
% 
% 
% end

%% PLOT Since Start
%  POOL shadow SHARES
%           vs COMPONENTS
           PLOT_RapFlag(  :,1) = POOL_return(:,1);                                    % Return Daily
           PLOT_RapFlag(  :,2) = cumprod(1+PLOT_RapFlag(:,1));                        % Return Cumulative
for t=2:T; PLOT_RapFlag(  t,3)=PLOT_RapFlag(t,2)^(365/t);                             % Return Annualized
           PLOT_RapFlag(  t,4)=std(PLOT_RapFlag(1:t,1));                              % StDev per day
           PLOT_RapFlag(  t,5)=sqrt(365)*PLOT_RapFlag(t,4);                           % StDev annualized
           PLOT_RapFlag(  t,6)=-maxdrawdown(PLOT_RapFlag(1:t,2));                     % DRAWDOWN
           PLOT_RapFlag(  t,7)=sqrt(365)*mean(PLOT_RapFlag(1:t,1))/PLOT_RapFlag(t,4); % SHARPE
end

for t=1:T
    PLOT_RapFlag_income( t,1)=sum(sum(POOL_dollars(1:t,[03 05 10 14 18]),1),2);
    PLOT_RapFlag_capgain(t,1)=sum(sum(POOL_dollars(1:t,[06 07 08 09 11 12 13 15 17 19]),1),2);
    PLOT_RapFlag_admin(  t,1)=sum(sum(POOL_dollars(1:t, 16),1),2);
    PLOT_RapFlag_return( t,1)=sum(sum(POOL_dollars(1:t,[03,05:19]),1),2);
    
    PLOT_RapFlag_income( t,2)=PLOT_RapFlag_income( t,1) / PLOT_RapFlag_return( t,1) * (PLOT_RapFlag(t,2)-1);
    PLOT_RapFlag_capgain(t,2)=PLOT_RapFlag_capgain(t,1) / PLOT_RapFlag_return( t,1) * (PLOT_RapFlag(t,2)-1);
    PLOT_RapFlag_admin(  t,2)=PLOT_RapFlag_admin(  t,1) / PLOT_RapFlag_return( t,1) * (PLOT_RapFlag(t,2)-1);
end

PLOT_component(:,:,1)=PLOT_RapFlag_income;  PLOT_component(1:365,2,1)=0;     PLOT_component(1897:1917,2,1)=0.9746; % Manual edit to fix plot.
% PLOT_component(:,:,2)=PLOT_RapFlag_capgain; PLOT_component(1:365,2,2)=0;
PLOT_component(:,:,2)=PLOT_RapFlag_admin;   PLOT_component(1:365,2,2)=0;

TITLE     = {['\color{black}','\fontsize{16}','\fontname{Roboto}',datestr(DATES(1,1)),' to ',datestr(DATES(end,1))]};

LEGEND    = {'Total Return';'Taxable Income';'Admin Costs'};
YLABEL    = 'RoE (%)';

figure
fun_PLOT_1_components(DATES(:,1),...
                      PLOT_RapFlag,...
                      PLOT_component,...
                      TITLE,LEGEND,YLABEL);
clear TITLE LEGEND


%% DATASTREAM \ BENCHMARKS
DSWS=datastreamws('ZUIO001','ORBIT592');

Bench_tickr  = {'MSACWF$'       ;'@BND'       ;'I$ALHYL'          ;'IBTRYAL'             ;'U:GLD'       ;'HFRIAWC'         };
Bench_field  = {'MSRI'          ;'RI'         ;'RI'               ;'RI'                  ;'RI'          ;'RI'              };
Bench_legnd  = {'World Equities';'World Bonds';'Global Junk Bonds';'USA Government Bonds';'Gold Bullion';'Hedge Fund Index'};

B=length(Bench_tickr);

% DS dates
for b=1:B
    data = history(DSWS,Bench_tickr(b),Bench_field(b),datestr(DATES(1,1)-1,'mm/DD/YYYY'),datestr(DATES(end,1),'mm/DD/YYYY'));
    data=timetable2table(data);
             DS_DATE=table2array(data(:,1));
             DS_RI  =table2array(data(:,2));
    for n=1:length(DS_RI)
        DS_data{b}(  n,1)=datenum(DS_DATE(n));                     % DATE
        DS_data{b}(  n,2)=          DS_RI(n);                      % RI
        DS_data{b}(1:n,3)=fun_P_to_R(DS_data{b}(1:n,2),'PERCENT'); % RETURN
    end
    DS_data{b}(1,:)=[];
    clear data
end
clear DSWS DS_RI DS_DATE

% 50-50 EQ-FI portfolio
DS_data{B+1}(:,1)=DS_data{1}(:,1);                          % DATES
DS_data{B+1}(:,2)=1;                                        % RI     (false for shortcut since not used)
DS_data{B+1}(:,3)=0.5*DS_data{1}(:,3)+0.5*DS_data{2}(:,3);  % RETURN (daily rebalancing)

Bench_tickr{B+1}={''};
Bench_field{B+1}={''};
Bench_legnd{B+1}={'Stocks+Bonds 50-50%'};

% Recount B = # of Benchmarks
B=length(DS_data);

% Insert to 365 day
PLOT_bench=zeros(T,8,B);
for b=1:B
    for t=2:T
        ind=find(DS_data{1,b}(:,1)==DATES(t,1));
        if ~isempty(ind); PLOT_bench(  t,1,b)=DS_data{1,b}(ind,3); end                                % Return Daily
                          PLOT_bench(1:t,2,b)=cumprod(1+PLOT_bench(1:t,1,b));                         % Return Cumulative
                          PLOT_bench(  t,3,b)=PLOT_bench(t,2,b)^(365/t);                              % Return Annualized
                          PLOT_bench(  t,4,b)=std(      PLOT_bench(1:t,1,b));                         % StDev per day
                          PLOT_bench(  t,5,b)=sqrt(365)*PLOT_bench(t,4,b);                            % StDev annualized
                          PLOT_bench(  t,6,b)=-maxdrawdown(PLOT_bench(1:t,2,b));                      % DRAWDOWN
                          PLOT_bench(  t,7,b)=sqrt(365)*mean(PLOT_bench(1:t,1,b))/PLOT_bench(t,4,b);  % SHARPE
             REGR=regress(POOL_return(1:t,1),[ones(t,1),PLOT_bench(1:t,1,b)]);
                          PLOT_bench(t,8,b)=((1+REGR(1))^365-1);                                      % ALPHA
                          PLOT_bench(t,9,b)=REGR(2);                                                  % BETA
                          clear REGR
    end
end

%% PLOT YTD
%  POOL shadow SHARES
%                vs
%              BENCHs
figure
% all YEARS
y=Y;
    % DATES
    if y==1;       DATES_Y = DATES(         1:EoY(y),1); end
    if y>=2;       DATES_Y = DATES(EoY(y-1)+1:EoY(y),1); end

    % RapFlag
    if y==1;       PLOT_RapFlag_Y(:,1) = POOL_return(         1:EoY(y),1);     end                  % Return Daily
    if y>=2;       PLOT_RapFlag_Y(:,1) = POOL_return(EoY(y-1)+1:EoY(y),1);     end                  % Return Daily
                   PLOT_RapFlag_Y(:,2) = cumprod(1+PLOT_RapFlag_Y(:,1));                            % Return Cumulative
    for d=2:length(PLOT_RapFlag_Y)
                   PLOT_RapFlag_Y(d,3) = PLOT_RapFlag_Y(d,2)^(365/t);                               % Return Annualized
                   PLOT_RapFlag_Y(d,4) = std(PLOT_RapFlag_Y(1:d,1));                                % StDev per day
                   PLOT_RapFlag_Y(d,5) = sqrt(365)*PLOT_RapFlag_Y(d,4);                             % StDev annualized
                   PLOT_RapFlag_Y(d,6) =-maxdrawdown(   PLOT_RapFlag_Y(1:d,2));                     % DRAWDOWN
                   PLOT_RapFlag_Y(d,7) = sqrt(365)*mean(PLOT_RapFlag_Y(1:d,1))/PLOT_RapFlag_Y(d,4); % SHARPE
    end
    % Benchmarks
    for b=1:B
        if y==1;       PLOT_bench_Y(:,1,b)=PLOT_bench(         1:EoY(y),1,b); end                         % Return Daily
        if y>=2;       PLOT_bench_Y(:,1,b)=PLOT_bench(EoY(y-1)+1:EoY(y),1,b); end                         % Return Daily
                       PLOT_bench_Y(:,2,b) = cumprod(1+PLOT_bench_Y(:,1,b));                              % Return Cumulative
        for d=2:length(PLOT_bench_Y)
                       PLOT_bench_Y(d,3,b) = PLOT_bench_Y(d,2,b)^(365/t);                               % Return Annualized
                       PLOT_bench_Y(d,4,b) = std(PLOT_bench_Y(1:d,1,b));                                % StDev per day
                       PLOT_bench_Y(d,5,b) = sqrt(365)*PLOT_bench_Y(d,4,b);                             % StDev annualized
                       PLOT_bench_Y(d,6,b) =-maxdrawdown(PLOT_bench_Y(1:d,2,b));                        % DRAWDOWN
                       PLOT_bench_Y(d,7,b) = sqrt(365)*mean(PLOT_bench_Y(1:d,1,b))/PLOT_bench_Y(d,4,b); % SHARPE
                       REGR=regress(PLOT_RapFlag_Y(1:d,1),[ones(d,1),PLOT_bench_Y(1:d,1,b)]);
                       PLOT_bench_Y(d,8,b)=((1+REGR(1))^365-1);                                         % ALPHA
                       PLOT_bench_Y(d,9,b)=    REGR(2);                                                 % BETA
        end           
    end

TITLE  =['\color{black}',num2str(year(DATES_Y(end))),': Comparables'];
LEGEND =['Rapaport Flagship', Bench_legnd{[1,7,2]}];

fun_PLOT_2(DATES_Y,...
               PLOT_RapFlag_Y,...
               PLOT_bench_Y(:,:,[1,7,2]),...
               TITLE,LEGEND,YLABEL);

%% PLOT YTD
%  POOL shadow SHARES
%                vs
%            COMPONENTS
figure
% all YEARS
y=Y;
    % DATES
    if y==1;       DATES_Y = DATES(         1:EoY(y),1); end
    if y>=2;       DATES_Y = DATES(EoY(y-1)+1:EoY(y),1); end

    % RapFlag
    if y==1;       PLOT_RapFlag_Y(:,1) = POOL_return(         1:EoY(y),1);     end                     % Return Daily
    if y>=2;       PLOT_RapFlag_Y(:,1) = POOL_return(EoY(y-1)+1:EoY(y),1);     end                     % Return Daily
                   PLOT_RapFlag_Y(:,2) = cumprod(1+PLOT_RapFlag_Y(:,1));                               % Return Cumulative
    for d=2:length(PLOT_RapFlag_Y)
                   PLOT_RapFlag_Y(  d,3) = PLOT_RapFlag_Y(d,2)^(365/t);                       % Return Annualized
                   PLOT_RapFlag_Y(  d,4) = std(           PLOT_RapFlag_Y(1:d,1));                      % StDev per day
                   PLOT_RapFlag_Y(  d,5) = sqrt(365)*     PLOT_RapFlag_Y(d,4);                         % StDev annualized
                   PLOT_RapFlag_Y(  d,6) =-maxdrawdown(   PLOT_RapFlag_Y(1:d,2));                      % DRAWDOWN
                   PLOT_RapFlag_Y(  d,7) = sqrt(365)*mean(PLOT_RapFlag_Y(1:d,1))/PLOT_RapFlag_Y(d,4);  % SHARPE
    end
    % Components
    if y==1;       POOL_dollars_Y=POOL_dollars(         1:EoY(y),:); end
    if y>=2;       POOL_dollars_Y=POOL_dollars(EoY(y-1)+1:EoY(y),:); end
    for d=1:length(POOL_dollars_Y)
        PLOT_RapFlag_income_Y( d,1)=sum(sum(POOL_dollars_Y(1:d,[03 05 10 14 18]),1),2);
        PLOT_RapFlag_capgain_Y(d,1)=sum(sum(POOL_dollars_Y(1:d,[06 07 08 09 11 12 13 15 17 19]),1),2);
        PLOT_RapFlag_admin_Y(  d,1)=sum(sum(POOL_dollars_Y(1:d,[16]),1),2)       ;
        PLOT_RapFlag_return_Y( d,1)=sum(sum(POOL_dollars_Y(1:d,[03,05:19]),1),2)        ;
        PLOT_RapFlag_income_Y( d,2)=PLOT_RapFlag_income_Y( d,1) / PLOT_RapFlag_return_Y( d,1) * (PLOT_RapFlag_Y(d,2)-1);
        PLOT_RapFlag_capgain_Y(d,2)=PLOT_RapFlag_capgain_Y(d,1) / PLOT_RapFlag_return_Y( d,1) * (PLOT_RapFlag_Y(d,2)-1);
        PLOT_RapFlag_admin_Y(  d,2)=PLOT_RapFlag_admin_Y(  d,1) / PLOT_RapFlag_return_Y( d,1) * (PLOT_RapFlag_Y(d,2)-1);
    end

PLOT_component_Y(:,:,1)=PLOT_RapFlag_income_Y;
PLOT_component_Y(:,:,2)=PLOT_RapFlag_admin_Y;
PLOT_component_Y(:,:,3)=PLOT_RapFlag_capgain_Y;

TITLE=['\color{black}',num2str(year(DATES_Y(end))),': Components'];
LEGEND    = {'Total Return';'Taxable Income';'Admin Costs';'Capital Gain'};

fun_PLOT_3(DATES_Y,...
               PLOT_RapFlag_Y,...
               PLOT_component_Y,...
               TITLE,LEGEND,YLABEL);

%% PLOT Since Start
%  POOL shadow SHARES
%           vs COMPARABLES

TITLE     = {['\color{black}','\fontsize{16}','\fontname{Roboto}',datestr(DATES(1,1)),' to ',datestr(DATES(end,1))]};
LEGEND    = ['RapFlag Total Return',Bench_legnd{[1,7,2]}];
YLABEL    = 'RoE (%)';

figure
fun_PLOT_1_comparables(DATES(:,1),...
                       PLOT_RapFlag,...
                       PLOT_bench(:,:,[1,7,2]),...
                       TITLE,LEGEND,YLABEL);
clear TITLE LEGEND

%% TABLE
%  PERFORMANCE MEASURES
%  POOL shadow Shares
import mlreportgen.dom.*;
filename= strcat('Performance_Measures_',datestr(DATES(end,1),'DDmmmYYYY'));
template= strcat('reports\templates\','POOL_PerformanceMeasures.dotx');
output  = strcat('reports\',filename);
DocObj  = Document(output,'docx',template);

    TABLE(1,1)=numberFormatter( POOL_prices(EoM(end),2)/POOL_prices(EoM(end-1),2)-1,'##.#%'); % Month's Return
    TABLE(2,1)=numberFormatter(PLOT_RapFlag_Y(end,2)-1,'##.#%');                                % YTD Return
    TABLE(3,1)=numberFormatter(  PLOT_RapFlag(end,2)-1,'##%');                                % Since Inception (cumulative)
    TABLE(4,1)=numberFormatter(  PLOT_RapFlag(end,3)-1,'##.#%');                              % Since Inception (annuazlied)
    TABLE(5,1)=numberFormatter(  PLOT_RapFlag(end,5),'##.#%');                                % Standard Deviation (annualized)
    TABLE(6,1)=numberFormatter(  PLOT_RapFlag(end,6),'##%');                                  % Worst Drawdown
    TABLE(7,1)=numberFormatter(  PLOT_RapFlag(end,7),'#.#');                                  % Sharpe Ratio
    TABLE(8,1)=cellstr('--');
    TABLE(9,1)=cellstr('--');
    
    for b=2:B+1
    temp=cumprod(1+PLOT_bench(EoM(end-1)+1:EoM(end),1,b-1))-1;    % Month's Return
    TABLE(1,b)=numberFormatter(temp(end),'##.#%');                % Month's Return
    TABLE(2,b)=numberFormatter(PLOT_bench_Y(end,2,b-1)-1,'##.#%');  % YTD Return
    TABLE(3,b)=numberFormatter(  PLOT_bench(end,2,b-1)-1,'##%');  % Since Inception (cumulative)
    TABLE(4,b)=numberFormatter(  PLOT_bench(end,3,b-1)-1,'##.#%');% Since Inception (annuazlied)
    TABLE(5,b)=numberFormatter(  PLOT_bench(end,5,b-1),'##.#%');  % Standard Deviation (annualized)
    TABLE(6,b)=numberFormatter(  PLOT_bench(end,6,b-1),'##%');    % Worst Drawdown
    TABLE(7,b)=numberFormatter(  PLOT_bench(end,7,b-1),'#.#');    % Sharpe Ratio
    TABLE(8,b)=numberFormatter(  PLOT_bench(end,8,b-1),'##%');    % Alpha
    TABLE(9,b)=numberFormatter(  PLOT_bench(end,9,b-1),'#.#');    % Beta
    end
    
    RowNames{1}='Month Return';
    RowNames{2}='Year to Date';
    RowNames{3}='Return Since Inception';
    RowNames{4}='Return Since Inception (annuzlized)';
    RowNames{5}='Standard Deviation (annualized)';
    RowNames{6}='Worst Drawdown';
    RowNames{7}='Sharpe Ratio (annualized)';
    RowNames{8}='Alpha over Benchmark (per year)';
    RowNames{9}='Beta  with Benchmark';

for b=1:B
    LEGEND_nospace{b,1}=Bench_legnd{b}(~isspace(Bench_legnd{b})); % Remove spaces
end
    LEGEND_nospace=['RapFlag';LEGEND_nospace];                    % Add RapFlag

TABLE=[LEGEND_nospace';TABLE];
TABLE=[[{''};RowNames'],TABLE];

moveToNextHole(DocObj);
while ~strcmp(DocObj.CurrentHoleId,'#end#')
          switch DocObj.CurrentHoleId
              case 'DATE';             value=datestr(DATES(end,1),'DD-mmm-YYYY'); append(DocObj,value);
              case 'TABLE';            value=TABLE;                               append(DocObj,value);
          end
          moveToNextHole(DocObj);
end
close(DocObj);
clear filename template output DocObj RowNames TABLE LEGEND_nospace


%% POOL
%  SERIES % in AA

import mlreportgen.dom.*;
filename = strcat('AA_',datestr(DATES(END,1),'DDmmmYYYY'));
template = strcat('reports\templates\','POOL_percent_AA.dotx');
output   = strcat('reports\',filename);
DocObj   = Document(output,'docx',template);

moveToNextHole(DocObj);
value=cellstr(datestr(DATES(END,1),'DD-mmm-YYYY'));
append(DocObj,value{:});
moveToNextHole(DocObj);
for aa=1:AA
    
    RowName{1,1}=['(S00) ',MNGR_name];
    FormatSpec='(S%02d) %s';
for s=2:S+1
    RowName{s,1}=sprintf(FormatSpec,s-1,CLTS_Names{s-1});
end
    RowName{S+2,1}='Total................................';
    
    CONTENT{1,1}=numberFormatter(MNGR_dollars_AA(END,end,aa),'$###,###');
    CONTENT{1,2}=numberFormatter(MNGR_percent_AA(END,end,aa),'##.####%');
    
    for s=2:S+1
    CONTENT{s,1}=numberFormatter(CLTS_dollars_AA(END,end,s-1,aa),'$###,###');
    CONTENT{s,2}=numberFormatter(CLTS_percent_AA(END,end,s-1,aa),'##.####%');
    end
    
    CONTENT{S+2,1}=numberFormatter(MNGR_dollars_AA(END,end,aa)+sum(CLTS_dollars_AA(END,end,:,aa),3),'$###,###');
    CONTENT{S+2,2}=numberFormatter(MNGR_percent_AA(END,end,aa)+sum(CLTS_percent_AA(END,end,:,aa),3),'##.####%');

    ColName{1}='DollarValue';
    ColName{2}='Percent';

TABLE = cell2table(CONTENT,'RowNames',RowName);
TABLE.Properties.VariableNames = ColName;

if strcmp(DocObj.CurrentHoleId,strcat('AA'   ,num2str(aa)));  value=num2str(aa);                                  append(DocObj,value);    moveToNextHole(DocObj); end
if strcmp(DocObj.CurrentHoleId,strcat('NAME' ,num2str(aa)));  value=POOL_names_AA{aa};                            append(DocObj,value);    moveToNextHole(DocObj); end
if strcmp(DocObj.CurrentHoleId,strcat('TYPE' ,num2str(aa)));  value=POOL_types_AA{aa};                            append(DocObj,value);    moveToNextHole(DocObj); end    
if strcmp(DocObj.CurrentHoleId,strcat('TABLE',num2str(aa)));  value=TABLE;                                        append(DocObj,value);    moveToNextHole(DocObj); end 

clear RowName CONTENT TABLE ColName
end

close(DocObj);
PDFfile=strcat(fullfile(pwd,'reports\',filename));
DOCfile=strcat(pwd,filesep,'reports\',filename,'.docx');
doc2pdf(DOCfile,PDFfile);
pause(0.1);
eval_expression=strcat('delete reports\',filename,'.docx');
eval(eval_expression);

clear filename template output DocObj PDFfile DOCfile 
