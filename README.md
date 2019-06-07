# Мониторинг очереди печати

## Описание
Программа создавалась для салона печати. 

Сначала генерируется номер заказа. После, когда в очередь добавляется новое задание печати, 
считывается количество страниц и копий у этого задания, и эти значения используются 
для подсчета общего количества страниц, при этом каждое задание приостанавливается.

Когда пользователь нажимает на кнопку "Готово", все задания возобновляются и в очередь добавляется файл "orderInfo". 
В этом файле находится информация о дате, времени, номере и количестве страниц выполненного заказа.

Помимо этого происходит логирование вышеуказанных данных в папку ```C:\Users\<USERNAME>\Documents\PrintOrders```

## Пример работы программы

![PrintOrders_2019-06-06_00-15-07](https://user-images.githubusercontent.com/25798995/58992283-0dcf4180-87f3-11e9-8c97-83527b77cb03.png)