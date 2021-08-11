## Проект receipts
Создание pdf-квитанций на основе excel-таблицы
### Для создания квитанций:
- Отредактируйте excel-файл "people.xlsx"
  ![alt text](screenshots/people.png "")
- В файле "people.xlsx" на листе "Настройки" укажите необходимые значения.
   Не забудьте сохранить изменения после редактирования файла
  ![alt text](screenshots/settings.png)
- Запустите файл ```main.py```.
- Если всё хорошо, то в папке должен появиться pdf-файл "Квитанции"
  ![alt text](screenshots/receipt.png)
- Если возникнет ошибка, то в открывшемся окне терминала появится описание ошибки
### Для создания исполняемого файла:
- Активируйте виртуальное окружение проекта:  
  ```source *путь к проекту*/venv/Scripts/activate```
- Запустите pyinstaller:  
```pyinstaller main.spec```
- В папке ```dist``` появится exe-файл "Создать квитанции". Поместите в ту же папку 
  файл "people.xlsx" и запустите программу
***
### Об авторе  
Брюшинин Алексей  
<megalaren@mail.ru>
