# DVB-S/S2 Satellite TV

Базовый Windows 10 проект для будущего приема DVB-S/S2 через BDA-совместимую карту, ориентирован на Prof Revolution 7300/7301 и Hotbird 13E.

## Текущий статус

- Приложение запускается как WPF UI.
- Реального DVB-устройства пока нет, используется `FakeDvbDevice`.
- Есть список стартовых транспондеров Hotbird 13E.
- Есть ручное добавление транспондера.
- Есть отмена сканирования, прогресс и статус текущего транспондера.
- Есть BDA/DirectShow detector для поиска DVB/AVStream фильтров в Windows.
- Есть симуляция сканирования и найденных FTA-каналов.
- При старте загружаются ранее сохраненные каналы.
- Найденные каналы сохраняются в `%LOCALAPPDATA%\DvbSatelliteTv\channels-hotbird-13e.json`.
- Рабочий список транспондеров сохраняется в `%LOCALAPPDATA%\DvbSatelliteTv\hotbird-13e-transponders.json`.
- Окно просмотра пока заглушка. libVLC будет подключен после появления реального transport stream path.

## Проекты

- `DvbSatelliteTv.App` - WPF интерфейс.
- `DvbSatelliteTv.Core` - модели, интерфейсы и дефолты Hotbird.
- `DvbSatelliteTv.Device` - слой DVB-устройства, сейчас симулятор.
- `DvbSatelliteTv.Device\BdaDeviceDetector.cs` - перечисление BDA Network Tuner, BDA Receiver/TS и DVB-like Capture фильтров.
- `DvbSatelliteTv.Storage` - локальное хранение каналов.
- `DvbSatelliteTv.App\Data\hotbird-13e-transponders.json` - стартовая офлайн-база транспондеров.

## Следующие этапы

1. Построить BDA tuning graph поверх найденных фильтров.
2. Заменить `FakeDvbDevice` на `BdaDvbDevice` после появления карты и драйвера.
3. Добавить настройку частоты, SR, поляризации и LNB 9750/10600.
4. Добавить чтение MPEG-TS и парсинг PAT/PMT/SDT.
5. Подключить libVLC для встроенного просмотра FTA-каналов.

## Команды

```powershell
dotnet build .\DvbSatelliteTv.slnx
dotnet run --project .\DvbSatelliteTv.App\DvbSatelliteTv.App.csproj
```
