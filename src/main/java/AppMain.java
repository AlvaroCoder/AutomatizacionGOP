import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.net.URL;
import java.util.Arrays;
import java.util.List;
import Scripts.ActualizaConta;
import nombradas.AgruparNombradas;
import cuadrillas.ControladorCuadrillas;
import teucost.controladores.ControladorCostoTeu;

public class AppMain {

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            // Look & Feel Nimbus (compatible con Java 8)
            try {
                for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                    if ("Nimbus".equals(info.getName())) {
                        UIManager.setLookAndFeel(info.getClassName());
                        break;
                    }
                }
            } catch (Exception ignored) { }

            DeckFrame frame = new DeckFrame();
            frame.setTitle("Panel Deck TPE");
            frame.setVisible(true);
        });
    }
}

class DeckFrame extends JFrame {

    // ---- UI Base ----
    private final JTextArea logArea = new JTextArea(8, 55);
    private final JProgressBar progressBar = new JProgressBar();

    // Botones (toolbar)
    private JButton btnConta;
    private JButton btnNombradas;
    private JButton btnTeuCost;
    private JButton btnCuadrillas;

    // Contenedor de formularios (CardLayout)
    private JPanel formContainer;
    private CardLayout formCards;

    // Formularios específicos
    private ContaFormPanel contaForm;
    private NombradasFormPanel nombradasForm;
    private ControladorCostoTeuFormPanel controladorForm;
    private CuadrillasFormPanel cuadrillasForm;

    // Paleta de azules suaves (texto negro)
    private final Color azulClaro1 = new Color(0xD7, 0xEA, 0xFE); // #D7EAFE
    private final Color azulClaro2 = new Color(0xB3, 0xDB, 0xFF); // #B3DBFF
    private final Color panelFondo = new Color(0xE9, 0xF4, 0xFF); // #E9F4FF

    DeckFrame() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(980, 850);
        setLocationRelativeTo(null);
        setResizable(true);

        // Ícono de la ventana (opcional, si existe en /icons/app.png)
        setIconImage(loadImageIconFromResource("/icons/app.png").getImage());

        setContentPane(buildContent());
    }

    private JPanel buildContent() {
        // Panel raíz con degradado azul claro
        GradientPanel root = new GradientPanel(azulClaro1, azulClaro2);
        root.setLayout(new BorderLayout(16, 15));
        root.setBorder(BorderFactory.createEmptyBorder(16, 16, 16, 16));

        // Encabezado + toolbar de botones (debajo del título)
        root.add(buildHeaderWithToolbar(), BorderLayout.NORTH);

        // Centro: Form dinámico
        root.add(buildCenterForms(), BorderLayout.CENTER);

        // Pie: barra de progreso + log
        root.add(buildFooter(), BorderLayout.SOUTH);

        return root;
    }

    // ------------------ HEADER + TOOLBAR ------------------
    private JComponent buildHeaderWithToolbar() {
        JPanel headerWrap = new JPanel();
        headerWrap.setOpaque(false);
        headerWrap.setLayout(new BoxLayout(headerWrap, BoxLayout.Y_AXIS));

        // ── Fila 1: Header con logo + título ──────────────────────────
        JPanel header = new JPanel() {
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                GradientPaint gp = new GradientPaint(0, 0, new Color(0xE6F2FF), getWidth(), 0, new Color(0xCCE6FF));
                g2.setPaint(gp);
                g2.fillRoundRect(0, 0, getWidth(), getHeight(), 24, 24);
                g2.dispose();
            }
        };
        header.setOpaque(false);
        header.setLayout(new BorderLayout());
        header.setMaximumSize(new Dimension(Integer.MAX_VALUE, 80));  // más compacto
        header.setBorder(BorderFactory.createEmptyBorder(8, 16, 8, 16));

        // Logo
        JLabel logo = new JLabel();
        ImageIcon logoIcon = loadLogoIcon();
        if (logoIcon != null && logoIcon.getIconHeight() > 0) {
            Image scaled = logoIcon.getImage().getScaledInstance(-1, 48, Image.SCALE_SMOOTH);
            logo.setIcon(new ImageIcon(scaled));
        }
        JPanel logoPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        logoPanel.setOpaque(false);
        logoPanel.add(logo);

        JLabel title = new JLabel("Panel Deck - GOP");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));
        title.setForeground(Color.BLACK);

        JLabel subtitle = new JLabel("Team Automatizaciones");
        subtitle.setFont(subtitle.getFont().deriveFont(Font.PLAIN, 12f));
        subtitle.setForeground(Color.BLACK);

        JPanel txt = new JPanel(new GridLayout(2, 1, 0, 2));
        txt.setOpaque(false);
        txt.add(title);
        txt.add(subtitle);

        JPanel leftHeader = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 4));
        leftHeader.setOpaque(false);
        leftHeader.add(logoPanel);
        leftHeader.add(txt);
        header.add(leftHeader, BorderLayout.WEST);

        JPanel toolbar = new JPanel(new GridLayout(1, 4, 8, 0));
        toolbar.setOpaque(false);
        toolbar.setMaximumSize(new Dimension(Integer.MAX_VALUE, 44));
        toolbar.setBorder(BorderFactory.createEmptyBorder(6, 0, 0, 0));

        btnConta = createActionButton("Actualiza Conta", "/icons/conta.png");
        btnConta.setToolTipText("Frecuencia: 2 veces por semana");
        btnConta.addActionListener(e -> showForm("CONTA"));

        btnNombradas = createActionButton("Agrupar Nombradas", "/icons/nombradas.png");
        btnNombradas.setToolTipText("Frecuencia: Ad-hoc");
        btnNombradas.addActionListener(e -> showForm("NOMBRADAS"));

        btnTeuCost = createActionButton("Costo x TEU", "/icons/costoteu.png");
        btnTeuCost.setToolTipText("Frecuencia de acuerdo a lo necesario");
        btnTeuCost.addActionListener(e -> showForm("CONTROLADOR"));

        btnCuadrillas = createActionButton("Cuadrillas SAP", "/icons/cuadrillas.png");
        btnCuadrillas.setToolTipText("Extrae datos de PDFs SAP a Excel");
        btnCuadrillas.addActionListener(e -> showForm("CUADRILLAS"));

        toolbar.add(btnConta);
        toolbar.add(btnNombradas);
        toolbar.add(btnTeuCost);
        toolbar.add(btnCuadrillas);

        headerWrap.add(header);
        headerWrap.add(toolbar);
        headerWrap.add(Box.createRigidArea(new Dimension(0, 10)));

        return headerWrap;
    }

    // ------------------ FORMULARIOS (CARD LAYOUT) ------------------
    private JComponent buildCenterForms() {
        formContainer = new JPanel();
        formCards = new CardLayout();

        formContainer.setLayout(formCards);
        formContainer.setOpaque(false);

        // Instanciamos los formularios con valores por defecto (tus hardcoded actuales)
        contaForm = new ContaFormPanel(
                "C:\\Users\\alvaro.pupuche\\Desktop\\PRACTICANTE PLANEAMIENTO\\THROUGHPUT GOP 2026.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\PRACTICANTE PLANEAMIENTO\\Conciliado Trafico 1Q 2026 semana 9 al 26 de Febreo 2026.xlsx",
                22, 90, 97
        );
        contaForm.setOnExecute(this::ejecutarActualizaConta);


        nombradasForm = new NombradasFormPanel(
                "C:\\Users\\alvaro.pupuche\\Desktop\\Nombradas COSMOS 2026\\Nombradas de Cosmos - Enero 2026",
                "C:\\Users\\alvaro.pupuche\\Desktop\\Nombradas COSMOS 2026"
        );
        nombradasForm.setOnExecute(this::ejecutarAgruparNombradas);

        controladorForm = new ControladorCostoTeuFormPanel(
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA THROUGHPUTS\\THROUGHPUT GOP 2026 3.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA NAVES\\DATA_NAVES_MARZO.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA CONCILIADOS\\Conciliado Trafico 1Q 2026 semana 13 al 31 de Marzo 2026.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA OMS\\OM- Mar26.xlsm",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\DATA MOVEHISTORY\\MoveHistory_ResumenMarzo.xlsx",
                "C:\\Users\\alvaro.pupuche\\Desktop\\TEUCOST\\RESUMEN EXCEL TEU COST.xlsx",
                "DatosNave",
                "MoveEvent"
        );
        controladorForm.setOnExecute(this::ejecutarControladorCostoTeu);

        cuadrillasForm = new CuadrillasFormPanel(
                "C:\\Temp\\CuadrillasTPE\\Naves",
                "C:\\Temp\\CuadrillasTPE\\resultado.xlsx"
        );
        cuadrillasForm.setOnExecute(this::ejecutarCuadrillas);
        formContainer.add(cuadrillasForm, "CUADRILLAS");

        formContainer.add(controladorForm, "CONTROLADOR");

        // Añadimos al CardLayout
        formContainer.add(contaForm, "CONTA");
        formContainer.add(nombradasForm, "NOMBRADAS");

        // Mostrar por defecto el primero
        formCards.show(formContainer, "CONTA");

        // Envolver en un panel con borde y título "Configuración"
        JPanel wrap = new JPanel(new BorderLayout());
        wrap.setOpaque(false);
        wrap.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(new Color(0xA7C7E7), 1, true),
                "Configuración del Script",
                TitledBorder.LEFT,
                TitledBorder.TOP
        ));

        ((TitledBorder) wrap.getBorder()).setTitleColor(Color.BLACK);

        wrap.setPreferredSize(new Dimension(0, 500));
        wrap.add(formContainer, BorderLayout.CENTER);

        return wrap;
    }

    private void showForm(String key) {
        formCards.show(formContainer, key);
        appendLog("Formulario activado: " + key);
    }

    // ------------------ FOOTER (PROGRESO + LOG) ------------------
    private JComponent buildFooter() {
        JPanel footer = new JPanel(new BorderLayout(8, 8));
        footer.setOpaque(false);

        progressBar.setStringPainted(true);
        progressBar.setForeground(new Color(0x2E7D32));
        progressBar.setBackground(Color.WHITE);
        progressBar.setPreferredSize(new Dimension(100, 26));

        logArea.setEditable(false);
        logArea.setForeground(Color.BLACK);
        logArea.setBackground(Color.WHITE);
        logArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        JScrollPane sp = new JScrollPane(logArea,
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        sp.setPreferredSize(new Dimension(100, 120));
        sp.setBorder(BorderFactory.createTitledBorder("Log de ejecución"));

        footer.add(progressBar, BorderLayout.NORTH);
        footer.add(sp, BorderLayout.CENTER);
        return footer;
    }

    // ------------------ Helpers de carga de iconos ------------------
    private ImageIcon loadImageIconFromResource(String pathInResources) {
        try {
            URL url = getClass().getResource(pathInResources);
            if (url == null) return new ImageIcon();
            ImageIcon icon = new ImageIcon(url);
            if (icon.getIconHeight() > 24) {
                Image img = icon.getImage().getScaledInstance(22, 22, Image.SCALE_SMOOTH);
                return new ImageIcon(img);
            }
            return icon;
        } catch (Exception e) {
            return new ImageIcon();
        }
    }

    private ImageIcon loadLogoIcon() {
        try {
            URL res = getClass().getResource("/assets/logo_TPE.jpg");
            if (res != null) return new ImageIcon(res);
        } catch (Exception ignored) { }
        File f = new File("main/java/assets/logo_TPE.jpg");
        if (f.exists()) {
            return new ImageIcon(f.getAbsolutePath());
        }
        return new ImageIcon(); // vacío
    }

    private JButton createActionButton(String text, String iconPath) {
        JButton btn = new JButton(text);
        btn.setFont(btn.getFont().deriveFont(Font.BOLD, 13f));
        btn.setForeground(Color.BLACK);
        btn.setFocusPainted(false);
        btn.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        btn.setHorizontalAlignment(SwingConstants.LEFT);
        btn.setPreferredSize(new Dimension(220, 42));

        ImageIcon icon = loadImageIconFromResource(iconPath);
        if (icon != null && icon.getIconWidth() > 0) {
            btn.setIcon(icon);
        }

        btn.setBackground(new Color(0xDFF0FF));
        btn.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(0x9CC6FF)),
                BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        btn.setOpaque(true);

        return btn;
    }

    // --- Notificaciones con sonido y JOptionPane ---
    private void notifyWithSound(String titulo, String mensaje, int messageType) {
        // 1) Sonido (intenta WAV; si no hay, hace beep)
        boolean sounded = playWavFromResources(
                messageType == JOptionPane.INFORMATION_MESSAGE
                        ? "/assets/success.wav"
                        : "/assets/error.wav"
        );
        if (!sounded) {
            Toolkit.getDefaultToolkit().beep();
        }

        SwingUtilities.invokeLater(() -> {
            JOptionPane.showMessageDialog(
                    this,
                    mensaje,
                    titulo,
                    messageType
            );
        });
    }

    /**
     * Intenta reproducir un WAV embebido en el classpath.
     * Retorna true si se pudo reproducir, false si falla (para usar beep).
     */
    private boolean playWavFromResources(String wavPath) {
        try (java.io.InputStream in = getClass().getResourceAsStream(wavPath)) {
            if (in == null) return false;
            try (javax.sound.sampled.AudioInputStream ais =
                         javax.sound.sampled.AudioSystem.getAudioInputStream(new java.io.BufferedInputStream(in))) {
                javax.sound.sampled.Clip clip = javax.sound.sampled.AudioSystem.getClip();
                clip.open(ais);
                clip.start();
                return true;
            }
        } catch (Exception ex) {
            // Si falla, devolvemos false para que haga beep
            return false;
        }
    }

    // ------------------ Ejecución con SwingWorker ------------------
    private void runTask(String nombre, Runnable tarea) {
        setButtonsEnabled(false);
        progressBar.setIndeterminate(true);
        progressBar.setString("Ejecutando: " + nombre);
        appendLog("▶ Iniciando: " + nombre);

        SwingWorker<Void, String> worker = new SwingWorker<Void, String>() {
            @Override
            protected Void doInBackground() {
                try {
                    tarea.run();
                    appendLog("✓ Finalizado correctamente: " + nombre);
                } catch (Exception ex) {
                    appendLog("✗ Error en " + nombre + ": " + ex.getMessage());

                    notifyWithSound("Hubo un problema",
                            "El proceso \"" + nombre + "\" terminó con errores.\nDetalle: " + ex.getMessage(),
                            JOptionPane.ERROR_MESSAGE);

                    ex.printStackTrace();
                }
                return null;
            }

            @Override
            protected void done() {
                progressBar.setIndeterminate(false);
                progressBar.setValue(100);
                progressBar.setString("Listo");
                setButtonsEnabled(true);

                notifyWithSound("Completado",
                        "El proceso \"" + nombre + "\" finalizó correctamente.",
                        JOptionPane.INFORMATION_MESSAGE);
            }
        };
        worker.execute();
    }


    private void setButtonsEnabled(boolean enabled) {
        if (btnConta != null)      btnConta.setEnabled(enabled);
        if (btnNombradas != null)  btnNombradas.setEnabled(enabled);
        if (btnTeuCost != null)    btnTeuCost.setEnabled(enabled);   // ← agregar esta línea

        if (contaForm != null)       contaForm.setExecuteEnabled(enabled);
        if (nombradasForm != null)   nombradasForm.setExecuteEnabled(enabled);
        if (controladorForm != null) controladorForm.setExecuteEnabled(enabled); // ← y esta
    }


    private void appendLog(String msg) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(msg + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    // ------------------ Acciones específicas (ahora usando valores del formulario) ------------------

    private void ejecutarActualizaConta() {
        final String pathThroughput = contaForm.getThroughputPath();
        final String pathConciliado = contaForm.getConciliadoPath();
        final String anchorObjetivo = contaForm.getObjetivoAnchor();      // nuevo
        final int rIni = contaForm.getRangoIni();
        final int rFin = contaForm.getRangoFin();

        runTask("Actualizar Conciliado Contabilidad", () -> {
            ActualizaConta actualizaConta = new ActualizaConta(pathThroughput, pathConciliado);
            actualizaConta.setObjetivoAnchor(anchorObjetivo);      // nuevo setter en tu clase
            actualizaConta.ejecutarRango(rIni, rFin);
        });
    }

    private void ejecutarAgruparNombradas() {
        final String carpetaEntrada = nombradasForm.getCarpetaEntrada();
        final String carpetaSalida = nombradasForm.getCarpetaSalida();

        runTask("Agrupar Nombradas", () -> {
            AgruparNombradas nombradas = new AgruparNombradas(carpetaEntrada, carpetaSalida);
            nombradas.procesar();
        });
    }

    private void ejecutarControladorCostoTeu() {
        final String rutaThroughput   = controladorForm.getRutaThroughput();
        final String rutaDataNaves    = controladorForm.getRutaDataNaves();
        final String rutaConciliado   = controladorForm.getRutaConciliado();
        final String rutaOMMensual    = controladorForm.getRutaOMMensual();
        final String rutaMoveHistory  = controladorForm.getRutaMoveHistory();
        final String rutaDestino      = controladorForm.getRutaExcelDestino();
        final String hojaCosmos       = controladorForm.getHojaCosmos();
        final String hojaMoveHistory  = controladorForm.getHojaMoveHistory();
        final int    numMes           = controladorForm.getNumMesOM();
        final List<String> listaNaves = controladorForm.getListaNaves();

        runTask("Controlador Costo x TEU", () -> {
            ControladorCostoTeu controlador = new ControladorCostoTeu(
                    rutaThroughput,
                    rutaDataNaves,
                    rutaConciliado,
                    rutaOMMensual,
                    rutaMoveHistory
            );
            controlador.setRutaExcelDestino(rutaDestino);
            controlador.setNombreHojaCosmos(hojaCosmos);
            controlador.setNombreHojaMoveHistory(hojaMoveHistory);
            controlador.setNumMesOM(numMes);
            controlador.escribirVariasVisitasEnTabla(listaNaves);
        });
    }

    // ---- Panel con degradado ----
    static class GradientPanel extends JPanel {
        private final Color c1;
        private final Color c2;

        GradientPanel(Color c1, Color c2) {
            this.c1 = c1;
            this.c2 = c2;
            setOpaque(false);
        }

        @Override
        protected void paintComponent(Graphics g) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            int w = getWidth(), h = getHeight();
            GradientPaint gp = new GradientPaint(0, 0, c1, 0, h, c2);
            g2.setPaint(gp);
            g2.fillRect(0, 0, w, h);
            g2.dispose();
            super.paintComponent(g);
        }
    }
    private void ejecutarCuadrillas() {
        final String rutaCarpetaNaves = cuadrillasForm.getRutaCarpetaNaves();
        final String rutaExcelDestino = cuadrillasForm.getRutaExcelDestino();

        runTask("Extracción Cuadrillas SAP", () -> {
            ControladorCuadrillas controlador = new ControladorCuadrillas(
                    rutaCarpetaNaves,
                    rutaExcelDestino
            );
            controlador.extraerDatosCuadrillas();
        });
    }
    // ------------------ Componentes reutilizables para formularios ------------------

    // Campo de archivo con etiqueta + JTextField + botón "Explorar..."
    static class FileField extends JPanel {
        private final JLabel label;
        private final JTextField text;
        private final JButton browse;
        private final boolean directoriesOnly;
        private final String fileDescription; // e.g. "Archivos Excel (*.xlsx)"

        FileField(String etiqueta, boolean directoriesOnly, String fileDescription) {
            setOpaque(true);
            setBackground(new Color(0xF7FBFF));
            setLayout(new GridBagLayout());
            this.directoriesOnly = directoriesOnly;
            this.fileDescription = fileDescription;

            label = new JLabel(etiqueta);
            label.setForeground(Color.BLACK);
            text = new JTextField(38);
            text.setForeground(Color.BLACK);

            browse = new JButton("Explorar...");
            browse.setBackground(new Color(0xDFF0FF));
            browse.setForeground(Color.BLACK);
            browse.setBorder(BorderFactory.createLineBorder(new Color(0x9CC6FF)));
            browse.setFocusPainted(false);
            browse.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            browse.addActionListener(e -> choose());

            GridBagConstraints gc = new GridBagConstraints();
            gc.insets = new Insets(4, 4, 4, 4);
            gc.gridx = 0; gc.gridy = 0; gc.anchor = GridBagConstraints.WEST;
            add(label, gc);

            gc.gridx = 1; gc.weightx = 1.0; gc.fill = GridBagConstraints.HORIZONTAL;
            add(text, gc);

            gc.gridx = 2; gc.weightx = 0; gc.fill = GridBagConstraints.NONE;
            add(browse, gc);
        }

        void setText(String value) { text.setText(value); }
        String getText() { return text.getText().trim(); }

        private void choose() {
            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle(directoriesOnly ? "Seleccionar carpeta" : "Seleccionar archivo");
            chooser.setFileSelectionMode(directoriesOnly ? JFileChooser.DIRECTORIES_ONLY : JFileChooser.FILES_ONLY);
            if (!directoriesOnly && fileDescription != null) {
                chooser.setAcceptAllFileFilterUsed(true);
                chooser.addChoosableFileFilter(new FileNameExtensionFilter(fileDescription, "xlsx", "xls"));
            }
            String current = text.getText().trim();
            if (!current.isEmpty()) {
                File base = new File(current);
                if (directoriesOnly && !base.isDirectory()) base = base.getParentFile();
                if (base != null && base.exists()) chooser.setCurrentDirectory(base);
            }
            int res = chooser.showOpenDialog(this);
            if (res == JFileChooser.APPROVE_OPTION) {
                File f = chooser.getSelectedFile();
                text.setText(f.getAbsolutePath());
            }
        }
    }

    static class IntField extends JPanel {
        private final JLabel label;
        private final JSpinner spinner;

        IntField(String etiqueta, int min, int max, int step, int initial) {
            setOpaque(false);
            setLayout(new FlowLayout(FlowLayout.LEFT, 4, 0));
            label = new JLabel(etiqueta);
            label.setForeground(Color.BLACK);
            spinner = new JSpinner(new SpinnerNumberModel(initial, min, max, step));
            add(label);
            add(spinner);
        }

        int get() { return (int) spinner.getValue(); }
        void set(int v) { spinner.setValue(v); }
        void setEnabledAll(boolean enabled) { label.setEnabled(enabled); spinner.setEnabled(enabled); }
    }

    static class BoolField extends JPanel {
        private final JCheckBox check;

        BoolField(String etiqueta, boolean initial) {
            setOpaque(false);
            setLayout(new FlowLayout(FlowLayout.LEFT, 4, 0));
            check = new JCheckBox(etiqueta, initial);
            check.setForeground(Color.BLACK);
            add(check);
        }

        boolean get() { return check.isSelected(); }
        void set(boolean v) { check.setSelected(v); }
        void setEnabledAll(boolean enabled) { check.setEnabled(enabled); }
    }

    // ------------------ Formularios concretos ------------------

    abstract static class BaseFormPanel extends JPanel {
        private final JButton ejecutarBtn = new JButton("Ejecutar");
        private Runnable onExecute;

        BaseFormPanel(String titulo, Color fondo) {
            setOpaque(true);
            setBackground(fondo);
            setLayout(new GridBagLayout());

            setBorder(BorderFactory.createTitledBorder(
                    BorderFactory.createLineBorder(new Color(0xA7C7E7), 1, true),
                    titulo,
                    TitledBorder.LEFT,
                    TitledBorder.TOP
            ));
            ((TitledBorder) getBorder()).setTitleColor(Color.BLACK);

            // Estilo botón ejecutar
            ejecutarBtn.setBackground(new Color(0xCFEAFF));
            ejecutarBtn.setForeground(Color.BLACK);
            ejecutarBtn.setFocusPainted(false);
            ejecutarBtn.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(new Color(0x9CC6FF)),
                    BorderFactory.createEmptyBorder(6, 12, 6, 12)
            ));
            ejecutarBtn.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            ejecutarBtn.addActionListener(e -> { if (onExecute != null) onExecute.run(); });
        }

        protected void addExecuteButton(GridBagConstraints gc) {
            gc.gridx = 0; gc.gridwidth = 3; gc.weightx = 0; gc.fill = GridBagConstraints.NONE;
            gc.insets = new Insets(10, 4, 4, 4);
            add(ejecutarBtn, gc);
        }

        void setOnExecute(Runnable r) { this.onExecute = r; }
        void setExecuteEnabled(boolean enabled) { ejecutarBtn.setEnabled(enabled); }
    }

// ---- Form: ActualizaConta ----
    static class ContaFormPanel extends BaseFormPanel {

        private final FileField tfThroughput;
        private final FileField tfConciliado;
        private final JTextField tfAnchorObjetivo;   // Ancla (Named Range o Hoja!Celda)
        private final IntField spRIni;
        private final IntField spRFin;

        // 'sheet' se mantiene en el constructor por compatibilidad, pero no se usa.
        ContaFormPanel(String pathThroughput, String pathConciliado, int sheet, int rIni, int rFin) {
            super("Actualiza Conciliado Contabilidad", new Color(0xF5FBFF));
            setPreferredSize(new Dimension(0, 420));

            GridBagConstraints gc = new GridBagConstraints();
            gc.insets = new Insets(6, 6, 6, 6);
            gc.anchor = GridBagConstraints.WEST;

            // --- Archivo Throughput ---
            tfThroughput = new FileField("Throughput (.xlsx):", false, "Archivos Excel (*.xlsx, *.xls)");
            tfThroughput.setText(pathThroughput);
            gc.gridx = 0; gc.gridy = 0; gc.gridwidth = 3; gc.fill = GridBagConstraints.HORIZONTAL; gc.weightx = 1.0;
            add(tfThroughput, gc);

            // --- Archivo Conciliado ---
            tfConciliado = new FileField("Conciliado Conta (.xlsx):", false, "Archivos Excel (*.xlsx, *.xls)");
            tfConciliado.setText(pathConciliado);
            gc.gridy = 1;
            add(tfConciliado, gc);

// --- Misma fila: Fila inicio + Fila final + Ancla objetivo ---
            spRIni = new IntField("Fila inicio:", 0, 10000, 1, rIni);
            spRFin = new IntField("Fila final:", 0, 10000, 1, rFin);

// subpanel para el nombre/ancla (label + textfield)
            JPanel pAnchorInline = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
            pAnchorInline.setOpaque(false);
            JLabel lAnchor = new JLabel("Hoja Objetivo Conta:");
            lAnchor.setForeground(Color.BLACK);
            tfAnchorObjetivo = new JTextField(18);
            tfAnchorObjetivo.setForeground(Color.BLACK);
            tfAnchorObjetivo.setToolTipText("Ejemplo: 'Cont-1-26' o 'HojaObjetivo!AT8'");
            tfAnchorObjetivo.setText("Cont-1-26"); // o "HojaObjetivo!AT8"
            pAnchorInline.add(lAnchor);
            pAnchorInline.add(tfAnchorObjetivo);

// contenedor de la fila combinada
            JPanel numbers = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 0));
            numbers.setOpaque(false);
            numbers.add(spRIni);
            numbers.add(spRFin);
            numbers.add(pAnchorInline);

            gc.gridy = 2;
            gc.gridwidth = 3;
            gc.fill = GridBagConstraints.NONE;
            gc.weightx = 0;
            add(numbers, gc);

// --- Botón ejecutar ---
            gc.gridy = 3;
            addExecuteButton(gc);

        }

        // Getters para usar en ejecutarActualizaConta()
        String getThroughputPath() { return tfThroughput.getText(); }
        String getConciliadoPath() { return tfConciliado.getText(); }
        String getObjetivoAnchor() { return tfAnchorObjetivo.getText().trim(); }
        int getRangoIni() { return spRIni.get(); }
        int getRangoFin() { return spRFin.get(); }
    }



    // ---- Form: AgruparNombradas ----
    static class NombradasFormPanel extends BaseFormPanel {
        private final FileField tfInDir;
        private final FileField tfOutDir;

        NombradasFormPanel(String inDir, String outDir) {
            super("Agrupar Nombradas", new Color(0xF5FBFF));
            GridBagConstraints gc = new GridBagConstraints();
            gc.insets = new Insets(6, 6, 6, 6);
            gc.anchor = GridBagConstraints.WEST;
            gc.gridx = 0; gc.gridy = 0; gc.gridwidth = 3; gc.fill = GridBagConstraints.HORIZONTAL; gc.weightx = 1.0;

            tfInDir = new FileField("Carpeta entrada:", true, null);
            tfInDir.setText(inDir);
            add(tfInDir, gc);

            tfOutDir = new FileField("Carpeta salida:", true, null);
            tfOutDir.setText(outDir);
            gc.gridy = 1;
            add(tfOutDir, gc);

            gc.gridy = 2;
            addExecuteButton(gc);
        }

        String getCarpetaEntrada() { return tfInDir.getText(); }
        String getCarpetaSalida() { return tfOutDir.getText(); }
    }

    static class ControladorCostoTeuFormPanel extends BaseFormPanel {
        private final FileField tfThroughput;
        private final FileField tfDataNaves;
        private final FileField tfConciliado;
        private final FileField tfOMMensual;
        private final FileField tfMoveHistory;
        private final FileField tfExcelDestino;
        private final JTextField tfHojaCosmos;
        private final JTextField tfHojaMoveHistory;
        private final IntField spNumMesOM;
        private final JTextField tfListaNaves;

        ControladorCostoTeuFormPanel(
                String rutaThroughput,
                String rutaControlCosmos,
                String rutaConciliado,
                String rutaOmMensual,
                String rutaMoveHistory,
                String rutaExcelDestino,
                String hojaCosmos,
                String hojaMoveHistory
        ) {
            super("Controlador Costo por TEU", new Color(0xF5FBFF));

            GridBagConstraints gc = new GridBagConstraints();
            gc.insets = new Insets(6, 6, 6, 6);
            gc.anchor = GridBagConstraints.WEST;
            gc.gridx = 0; gc.gridwidth = 3;
            gc.fill = GridBagConstraints.HORIZONTAL;
            gc.weightx = 1.0;

            // --- Archivos de entrada ---
            tfThroughput = new FileField("Throughput (.xlsx):", false, "Archivos Excel (*.xlsx, *.xls)");
            tfThroughput.setText(rutaThroughput);
            gc.gridy = 0; add(tfThroughput, gc);

            tfDataNaves = new FileField("Data Naves (.xlsx):", false, "Archivos Excel (*.xlsx, *.xls)");
            tfDataNaves.setText(rutaControlCosmos);
            gc.gridy = 1; add(tfDataNaves, gc);

            tfConciliado = new FileField("Conciliado (.xlsx):", false, "Archivos Excel (*.xlsx, *.xls)");
            tfConciliado.setText(rutaConciliado);
            gc.gridy = 2; add(tfConciliado, gc);

            tfOMMensual = new FileField("OM Mensual (.xlsm):", false, "Archivos Excel (*.xlsm, *.xlsx)");
            tfOMMensual.setText(rutaOmMensual);
            gc.gridy = 3; add(tfOMMensual, gc);

            tfMoveHistory = new FileField("Move History (.xlsx):", false, "Archivos Excel (*.xlsx, *.xls)");
            tfMoveHistory.setText(rutaMoveHistory);
            gc.gridy = 4; add(tfMoveHistory, gc);

            tfExcelDestino = new FileField("Excel Destino (.xlsx):", false, "Archivos Excel (*.xlsx, *.xls)");
            tfExcelDestino.setText(rutaExcelDestino);
            gc.gridy = 5; add(tfExcelDestino, gc);

            // --- Fila: Hoja Cosmos + Hoja MoveHistory + Mes OM ---
            JPanel rowConfig = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 0));
            rowConfig.setOpaque(false);

            // Hoja Cosmos
            JPanel pHojaCosmos = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
            pHojaCosmos.setOpaque(false);
            JLabel lHojaCosmos = new JLabel("Hoja Cosmos:");
            lHojaCosmos.setForeground(Color.BLACK);
            tfHojaCosmos = new JTextField(12);
            tfHojaCosmos.setForeground(Color.BLACK);
            tfHojaCosmos.setText(hojaCosmos);
            pHojaCosmos.add(lHojaCosmos);
            pHojaCosmos.add(tfHojaCosmos);

            // Hoja MoveHistory
            JPanel pHojaMH = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
            pHojaMH.setOpaque(false);
            JLabel lHojaMH = new JLabel("Hoja MoveHistory:");
            lHojaMH.setForeground(Color.BLACK);
            tfHojaMoveHistory = new JTextField(12);
            tfHojaMoveHistory.setForeground(Color.BLACK);
            tfHojaMoveHistory.setText(hojaMoveHistory);
            pHojaMH.add(lHojaMH);
            pHojaMH.add(tfHojaMoveHistory);

            // Mes OM
            spNumMesOM = new IntField("Mes OM (1-12):", 1, 12, 1, 3);

            rowConfig.add(pHojaCosmos);
            rowConfig.add(pHojaMH);
            rowConfig.add(spNumMesOM);

            gc.gridy = 6; gc.fill = GridBagConstraints.NONE; gc.weightx = 0;
            add(rowConfig, gc);

            // --- Lista de naves ---
            JPanel pNaves = new JPanel(new BorderLayout(6, 0));
            pNaves.setOpaque(false);
            JLabel lNaves = new JLabel("Lista de naves (separadas por coma):");
            lNaves.setForeground(Color.BLACK);
            tfListaNaves = new JTextField("171-26, 173-26, 175-26, 174-26, 176-26, 186-26, 178-26");
            tfListaNaves.setForeground(Color.BLACK);
            tfListaNaves.setToolTipText("Ejemplo: 171-26, 173-26, 175-26");
            pNaves.add(lNaves, BorderLayout.WEST);
            pNaves.add(tfListaNaves, BorderLayout.CENTER);

            gc.gridy = 7; gc.fill = GridBagConstraints.HORIZONTAL; gc.weightx = 1.0;
            add(pNaves, gc);

            // --- Botón ejecutar ---
            gc.gridy = 8;
            addExecuteButton(gc);
        }

        // Getters
        String getRutaThroughput()      { return tfThroughput.getText(); }
        String getRutaDataNaves()       { return tfDataNaves.getText(); }
        String getRutaConciliado()      { return tfConciliado.getText(); }
        String getRutaOMMensual()       { return tfOMMensual.getText(); }
        String getRutaMoveHistory()     { return tfMoveHistory.getText(); }
        String getRutaExcelDestino()    { return tfExcelDestino.getText(); }
        String getHojaCosmos()          { return tfHojaCosmos.getText().trim(); }
        String getHojaMoveHistory()     { return tfHojaMoveHistory.getText().trim(); }
        int    getNumMesOM()            { return spNumMesOM.get(); }

        List<String> getListaNaves() {
            String raw = tfListaNaves.getText().trim();
            if (raw.isEmpty()) return new java.util.ArrayList<>();
            return Arrays.stream(raw.split(","))
                    .map(String::trim)
                    .filter(s -> !s.isEmpty())
                    .collect(java.util.stream.Collectors.toList());
        }
    }
    static class CuadrillasFormPanel extends BaseFormPanel {
        private final FileField tfCarpetaNaves;
        private final FileField tfExcelDestino;
        private final JLabel lblTesseractEstado;

        CuadrillasFormPanel(String rutaCarpetaNaves, String rutaExcelDestino) {
            super("Extracción Cuadrillas SAP (OCR)", new Color(0xF5FBFF));

            GridBagConstraints gc = new GridBagConstraints();
            gc.insets = new Insets(6, 6, 6, 6);
            gc.anchor = GridBagConstraints.WEST;
            gc.gridx = 0; gc.gridwidth = 3;
            gc.fill = GridBagConstraints.HORIZONTAL;
            gc.weightx = 1.0;

            // --- Carpeta de naves (PDFs SAP origen) ---
            tfCarpetaNaves = new FileField("Carpeta PDFs SAP:", true, null);
            tfCarpetaNaves.setText(rutaCarpetaNaves);
            gc.gridy = 0;
            add(tfCarpetaNaves, gc);

            // --- Excel destino ---
            tfExcelDestino = new FileField("Excel destino (.xlsx):", false,
                    "Archivos Excel (*.xlsx)");
            tfExcelDestino.setText(rutaExcelDestino);
            gc.gridy = 1;
            add(tfExcelDestino, gc);

            // --- Info sobre carpeta temporal ---
            JLabel lblInfo = new JLabel(
                    "ℹ Los PDFs se copiarán temporalmente a C:\\Temp\\CuadrillasTPE\\ antes del OCR.");
            lblInfo.setForeground(new Color(0x555555));
            lblInfo.setFont(lblInfo.getFont().deriveFont(Font.ITALIC, 11f));
            gc.gridy = 2;
            gc.fill = GridBagConstraints.NONE;
            gc.weightx = 0;
            add(lblInfo, gc);

            // --- Estado de Tesseract (detectado automáticamente) ---
            String estadoTess = detectarTesseract();
            lblTesseractEstado = new JLabel("🔍 Tesseract: " + estadoTess);
            lblTesseractEstado.setForeground(
                    estadoTess.startsWith("✓") ? new Color(0x1B6B1B) : new Color(0xAA2200));
            lblTesseractEstado.setFont(lblTesseractEstado.getFont().deriveFont(Font.PLAIN, 11f));
            gc.gridy = 3;
            add(lblTesseractEstado, gc);

            // --- Botón ejecutar ---
            gc.gridy = 4;
            addExecuteButton(gc);
        }

        // Detecta si Tesseract está instalado en rutas conocidas
        private String detectarTesseract() {
            String[] rutas = {
                    "C:/Program Files/Tesseract-OCR/tessdata",
                    "C:/Program Files (x86)/Tesseract-OCR/tessdata",
                    System.getProperty("user.home") +
                            "/AppData/Local/Programs/Tesseract-OCR/tessdata"
            };
            for (String ruta : rutas) {
                if (new File(ruta).exists()) {
                    return "✓ Encontrado en " + ruta;
                }
            }
            return "✗ No encontrado — instalar desde github.com/UB-Mannheim/tesseract/wiki";
        }

        String getRutaCarpetaNaves() { return tfCarpetaNaves.getText(); }
        String getRutaExcelDestino() { return tfExcelDestino.getText(); }
    }
}