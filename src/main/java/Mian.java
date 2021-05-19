public class Mian {

    private static String wr1 = "Москва, осень 2016 года. В город возвращается 27-летний Илья Горюнов — бывший студент-филолог, отсидевший семь лет за подкинутые в клубе наркотики. Главный герой не узнаёт столицу, которая сильно изменилась за эти годы. В особенности его глаз цепляется за смартфоны, которые раньше были «только у продвинутых», а сейчас — у каждого.\n" +
            "Горюнов направляется домой, в подмосковную Лобню. Она, в отличие от Москвы, осталась той же. Героя охватывают воспоминания: школьные годы, бывшая девушка, друзья. Илья мечтает наконец увидеть мать, с которой он жил до ареста, но, приехав домой, узнаёт, что мать умерла от инфаркта за день до его возвращения.\n" +
            "Илья приглашает к себе старого друга Серёгу, но при встрече понимает, что они стали друг другу чужими. С помощью телефона друга Горюнов находит во «ВКонтакте» страницу Петра Хазина — лейтенанта ФСКН, который семь лет назад отправил Илью за решётку, подкинув наркотики. Сам герой называет его про себя «Сукой». Видя довольное лицо человека, сломавшего ему жизнь, герой решает отомстить.\n" +
            "В итоге в руках Горюнова оказывается смартфон его обидчика, в котором спрятана вся жизнь и весь компромат на Хазина. Понимая, что сам он обречён, Илья начинает жить жизнью другого человека, используя лишь его телефон. В дальнейшем герой пытается разрешить семейные конфликты Суки и влюбляется по видео и фото в телефоне Хазина в его девушку Нину. Притворяясь Хазиным, герой узнаёт всю его подноготную и находит способ получить 250 тысяч евро от его сообщников. В конце перед героем встаёт моральный выбор: покинуть страну с этими деньгами или спасти Нину.";

    private static String wr2 = "По словам автора, идея произведения вызревала несколько лет, тогда как сама работа заняла несколько месяцев. Глуховский демонстрировал рукопись силовикам и отсидевшим преступникам. По его словам, один из них сказал: «Вот прямо про меня написано»[5]. " +
            " 20 марта 2017 года в официальном сообществе во «ВКонтакте» появился первый вариант обложки будущего романа. Спустя десять дней на страницах сообщества начали публиковаться первые отрывки из «Текста», а также трейлеры и интервью Глуховского о книге. Затем в группе во «ВКонтакте» были опубликованы целиком несколько первых глав. Глуховский также устраивал чтения собственного романа в live-трансляциях. 14 июня автор презентовал роман в пресс-центре НСН в Москве, на следующий день стартовала продажа новой книги." +
            " У романа есть собственный саундтрек, опубликованный на странице произведения во «ВКонтакте». Роман опубликован в печатном и электронном виде, также доступна аудиокнига. ";

    public static void main(String[] args) {
        App app = new App();
        try {
            app.writeIntoExcel("excel.xlsx", wr1);
            String s = app.readFromExcel("excel.xlsx");
            app.writeIntoExcel("excel_replace.xlsx", wr2);
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }    }

}
