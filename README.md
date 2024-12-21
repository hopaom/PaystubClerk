# PaystubClerk
Program to generate payslip from payroll ledger
<br/>
<br/>
급여대장으로부터 급여명세서를 생성하는 프로그램
<br/>
<br/>
<img src="./manual.png" width="396" height="549"/>
<br/>

### 필요한 패키지 파일
openpyxl  
tkinter
<br/>
<br/>

### 파일 빌드
##### 콘솔을 관리자 권한으로 실행 해야 한다
`pyinstaller.exe --onefile --windowed --add-data "./manual.png;." --noconsole -n=PaystubClerk --icon=payst.ico paystubclerk.py`
