# Slidebrary

Корпоративная библиотека материалов в виде add-in для Microsoft PowerPoint.

## Установка для пользователя на macOS

Пользователю не нужны Git, Node.js или VS Code.

1. Откройте [manifest.xml в GitHub](https://github.com/gloomeries/Slide-library/blob/main/manifest.xml)
   и нажмите **Download raw file**.
2. Откройте Terminal.
3. Выполните команды:

   ```bash
   mkdir -p "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
   cp "$HOME/Downloads/manifest.xml" "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/"
   ```

4. Полностью закройте и снова откройте PowerPoint.
5. Откройте **Главная → Надстройки → Slidebrary**.

Если браузер сохранил файл под другим названием, например `manifest (1).xml`, переименуйте
его в `manifest.xml` перед выполнением команды `cp`.

## Личная библиотека

1. Откройте Slidebrary и нажмите на иконку пользователя.
2. Скачайте шаблон личной папки и распакуйте архив.
3. Разложите материалы по разделам внутри `Slidebrary Personal`.
4. Выберите корневую папку в личном кабинете Slidebrary и нажмите **Сохранить**.

Личные файлы читаются с устройства и не отправляются в GitHub или облако. После полного
закрытия PowerPoint папку потребуется выбрать повторно.

## Разработка

```bash
npm install
npm start
```

Локальная версия использует `https://localhost:3000`. Остановить тестовую сессию:

```bash
npm run stop
```

## Публикация

Push в ветку `main` автоматически запускает production-сборку и публикует папку `dist`
в GitHub Pages. Основной `manifest.xml` уже использует production-адрес:

```text
https://gloomeries.github.io/Slide-library/
```

Если GitHub Pages настраивается впервые, владелец репозитория должен один раз открыть
**Settings → Pages → Build and deployment → Source** и выбрать **GitHub Actions**.
