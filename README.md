# chang-gung-doctor-scheduling
這個程式是用來排 chang-gung 醫師的下個月班表  

# 需求
給填國定假日  
預約不值班日  

# 排班規則
特休一次休半個月，分成每個月1-15號, 16-月底  
平日點數 1 假日 2  
三天內同一人只能排一班  
一天兩線，第一線(U) R3, R4，第二線(U2) R4, 盡量R5  
星期五,六的二線班(U2)為同一人  

# 要填的資料 what you need to do
改member.xlsx  
Sheet1 國定假日往下每格填下個月國定假日日期  
Sheet2 人名下面第一個數字代表是R3, R4, R5，職級下面填下個月每人休假日期  
特休上半月填first half，下半月填second half，休一天的填日期數字  

# 執行 schedule.exe
執行schedule.exe會根據member.xlsx排下個月班表，班表會輸出到schedule.xlsx  
要重新排記得把schedule.xlsx關掉再重新執行schedule.exe，不然輸出會失敗  

# 注意
exe檔可能被windows defender認為是病毒移除，要暫時關掉windows defender或是加入白名單  
