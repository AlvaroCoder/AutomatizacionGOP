package postmortem.componentes;

import java.util.ArrayList;
import java.util.List;

public class ConstructorHtml {

    private final List<Componente> componentes = new ArrayList<>();

    public ConstructorHtml agregar(Componente componente) {
        componentes.add(componente);
        return this; // permite encadenar
    }

    public String build() {
        // Recopilar todo el CSS de los componentes
        StringBuilder cssTotal = new StringBuilder();
        cssTotal.append("* { box-sizing: border-box; margin: 0; padding: 0; "
                + "font-family: Arial, sans-serif; font-size: 11px; }"
                + "body { background: #f0f4f0; width: 1200px; }");
        for (Componente c : componentes) {
            cssTotal.append(c.css());
        }

        // Recopilar el HTML de los componentes
        StringBuilder htmlBody = new StringBuilder();
        for (Componente c : componentes) {
            htmlBody.append(c.render());
        }

        return "<!DOCTYPE html><html lang='es'><head>"
                + "<meta charset='UTF-8'/>"
                + "<title>POST-MORTEM</title>"
                + "<style>" + cssTotal + "</style>"
                + "</head><body>"
                + htmlBody
                + "</body></html>";
    }
}