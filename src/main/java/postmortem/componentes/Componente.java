package postmortem.componentes;

public interface Componente {
    String render();
    default String css() {return "";}
}
