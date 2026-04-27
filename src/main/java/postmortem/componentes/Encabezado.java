package postmortem.componentes;

public class Encabezado implements Componente{
    private final String logoPath;

    public Encabezado(String logoPath){
        this.logoPath = logoPath;
    }

    @Override
    public String render(){
        return "<div class='title-bar'>"
                + "<span class='back-arrow'>&#8592;</span>"
                + "<span class='title-text'>POST-MORTEM</span>"
                + "<img src='"+logoPath+"' class='title-logo' alth='LogoTpe' />"
                + "</div>";
    }

    @Override
    public String css(){
        return ".title-bar {" +
                "  background-color: black;" +
                "  padding: 8px;" +
                "  color: white;" +
                "  display: flex;" +
                "  flex-direction: row;" +
                "  justify-content: space-between;" +
                "  align-items: space-between;" +
                "}" +
                ".title-text {" +
                "  font-weight: bold;" +
                "  letter-spacing: 2px;" +
                "}" +
                ".title-logo { height: 36px; object-fit: contain; }";
    }
}
