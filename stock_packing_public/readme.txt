1. stock_id.txt 
2. income_statement 
存放所有股票的xml檔，若同個季度，則不再從web抓取。新季度時，請清空此資料匣檔案或另行打包出來。程式會從web抓新的xml回來存放。

3.執行 Quarterly_Analysis.exe
產生
error_id.txt => 解析有問題的股票代碼
營業收入淨額1043Q.txt => 104年第三季損益表中，營業收入淨額連續三季成長，且皆為非負數
營業毛利1043Q.txt => 104年第三季損益表中，營業毛利連續三季成長，且皆為非負數
營業利益1043Q.txt => 104年第三季損益表中，營業利益連續三季成長，且皆為非負數
每股盈餘1043Q.txt => 104年第三季損益表中，每股盈餘連續三季成長，且皆為非負數
四冠王.txt => 上面四項皆達成者

重新執行 Quarterly_Analysis.exe 時，請將output的.txt刪除

4.新增 Revenue.exe 月營收創新高快報
xml 檔存放在revenue 資料夾內
