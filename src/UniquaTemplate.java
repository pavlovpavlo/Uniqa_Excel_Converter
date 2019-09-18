public class UniquaTemplate
{
    private int article ;
    private String name="";
    private String price=null;
    private String discount="0%";
    private String uniqaprice= null;
    private String ndsflag="Нет";
    private double quantity=-1;

    public UniquaTemplate(){}

    public int getArticle() {
        return article;
    }

    public String getName() {
        return name;
    }

    public String getPrice() {
        return price;
    }

    public String getDiscount() {
        return discount;
    }

    public String getUniqaprice() {
        return uniqaprice;
    }

    public String getNdsflag() {
        return ndsflag;
    }

    public double getQuantity() {
        return quantity;
    }

    public void setArticle(int article) {
        this.article = article;
    }

    public void setName(String name) {
        this.name = name.replaceAll("[`'']","\"");
    }

    public void setPrice(String price) {
        this.price = price;
    }

    public void setDiscount(String discount) {
        this.discount = discount;
    }

    public void setUniqaprice(String uniqaprice) {
        this.uniqaprice = uniqaprice;
    }

    public void setNdsflag(String ndsflag) {
        this.ndsflag = ndsflag;
    }

    public void setQuantity(double quantity) {
        this.quantity = quantity;
    }
}
