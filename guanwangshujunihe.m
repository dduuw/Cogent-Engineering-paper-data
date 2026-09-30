%% ============================================================
%  稳健版：读取 B~H 列 (7个影响因素) -> 线性拟合 I 列
% ============================================================
clear; clc;

%% ---------- 参数 ----------
polyDegree     = 1;      % 1 = 线性; 2 = 二次多项式（含平方+交叉项）
useStandardize = false;  % 量纲差异大时改 true

dataFolder = uigetdir('C:\Users\USER\Desktop', '请选择 Excel 文件夹');
if isequal(dataFolder,0), error('未选择文件夹'); end

files = dir(fullfile(dataFolder, '*.xlsx'));
files = files(~contains({files.name}, 'Extract_Result') & ...
              ~contains({files.name}, 'Fit_'));

fprintf('找到 %d 个文件\n', numel(files));

X_all = [];  Y_all = [];  fileRecord = {};

for k = 1:numel(files)
    fp = fullfile(files(k).folder, files(k).name);
    try
        % ---- 读整张表（不用 Range，避免自动区域识别出错）----
        M = readmatrix(fp);          % 非数值单元格 -> NaN

        if isempty(M) || size(M,2) < 9
            warning('%s: 列数不足 9，跳过（size=[%d x %d]）', ...
                files(k).name, size(M,1), size(M,2));
            continue;
        end

        % ---- 按列号取：B=2,...,H=8 ; I=9 ----
        X = M(:, 2:8);      % B~H  (7 列)
        Y = M(:, 9);        % I    (第9列)

        % ---- 只保留 X、Y 全为有效数值的行 ----
        valid = all(~isnan(X),2) & ~isnan(Y);
        fprintf('[%d/%d] %s: 共 %d 行, 有效 %d 行\n', ...
            k, numel(files), files(k).name, size(M,1), sum(valid));

        if ~any(valid)
            warning('  -> 该文件无有效数据，跳过');
            continue;
        end

        X = X(valid,:);  Y = Y(valid);
        X_all = [X_all; X];
        Y_all = [Y_all; Y];
        fileRecord = [fileRecord; repmat({files(k).name}, size(X,1), 1)];

    catch ME
        warning('%s 读取失败: %s', files(k).name, ME.message);
    end
end

if isempty(X_all)
    error(['仍未收集到有效数据。常见原因：\n' ...
           '1) I 列不是纯数字（含文字/单位/公式未计算）\n' ...
           '2) B~H 列有空单元格\n' ...
           '3) 数据不在第一个 sheet\n' ...
           '请先打印 M 查看结构。']);
end

n = size(X_all,1);
fprintf('\n共收集样本 %d 条，自变量 7 个\n', n);

%% ---------- 标准化（可选）----------
if useStandardize
    [Xin, mu, sigma] = zscore(X_all);
else
    Xin = X_all;
end

%% ---------- 构造多项式特征 ----------
[Xf, featNames] = buildPolyFeatures(Xin, polyDegree);
p = size(Xf,2);
fprintf('特征数(含常数) = %d\n', p);

if n < p
    warning('样本数 < 特征数，建议减小 polyDegree 或收集更多数据。');
end

%% ---------- 最小二乘 ----------
b     = Xf \ Y_all;
Ypred = Xf * b;
resid = Y_all - Ypred;

SSE  = sum(resid.^2);
SST  = sum((Y_all - mean(Y_all)).^2);
R2   = 1 - SSE/SST;
R2adj= 1 - (1-R2)*(n-1)/(n-p);
RMSE = sqrt(SSE/n);

fprintf('\n===== 拟合结果 =====\n');
fprintf('R²     = %.4f\n调整R² = %.4f\nRMSE   = %.4f\n', R2, R2adj, RMSE);

fprintf('\n%-15s %16s\n','项','系数');
fprintf('%s\n', repmat('-',1,34));
for i = 1:p
    fprintf('%-15s %16.6g\n', featNames{i}, b(i));
end

%% ---------- 生成可读公式 ----------
formula = sprintf('I = %.6g', b(1));
for i = 2:p
    if b(i) >= 0
        formula = [formula, sprintf(' + %.6g*%s', b(i), featNames{i})]; %#ok<AGROW>
    else
        formula = [formula, sprintf(' - %.6g*%s', abs(b(i)), featNames{i})]; %#ok<AGROW>
    end
end
fprintf('\n拟合公式:\n%s\n', formula);

%% ---------- 保存 ----------
coefTbl = table(featNames(:), b, 'VariableNames', {'Term','Coefficient'});
predTbl = table(fileRecord, Y_all, Ypred, resid, ...
    'VariableNames', {'File','I_true','I_pred','Residual'});

out1 = fullfile(dataFolder,'Fit_Coefficients.xlsx');
out2 = fullfile(dataFolder,'Fit_Predictions.xlsx');
for f = {out1,out2}
    if exist(f{1},'file'), try delete(f{1}); catch, end, end
end
writetable(coefTbl, out1);
writetable(predTbl, out2);
fprintf('\n结果已保存:\n  %s\n  %s\n', out1, out2);

%% ---------- 画图 ----------
figure('Color','w');
scatter(Y_all, Ypred, 25, 'filled'); hold on;
lims = [min([Y_all;Ypred]), max([Y_all;Ypred])];
plot(lims, lims, 'r--','LineWidth',1.5);
xlabel('真实 I 列'); ylabel('预测 I 列');
title(sprintf('R^2=%.4f  RMSE=%.4f', R2, RMSE));
grid on; axis equal;

%% ============ 辅助函数 ============
function [Xf, names] = buildPolyFeatures(X, degree)
    [n, p] = size(X);
    Xf = ones(n,1);
    names = {'常数'};
    % 一阶项
    for i = 1:p
        Xf = [Xf, X(:,i)];                 %#ok<AGROW>
        names{end+1} = sprintf('x%d', i);  %#ok<AGROW>
    end
    % 二次项（可选）
    if degree >= 2
        for i = 1:p
            Xf = [Xf, X(:,i).^2];               %#ok<AGROW>
            names{end+1} = sprintf('x%d^2', i); %#ok<AGROW>
        end
        for i = 1:p-1
            for j = i+1:p
                Xf = [Xf, X(:,i).*X(:,j)];               %#ok<AGROW>
                names{end+1} = sprintf('x%d*x%d', i, j); %#ok<AGROW>
            end
        end
    end
end