@{
    # ========================================================================
    #  Themes.psd1
    #  Single source of truth for all GUI theme palettes.
    #
    #  Consumed by:
    #    - Invoke-PatchGUI.ps1            (reads preferred theme at launch)
    #    - Invoke-PatchGUI-Gallery.ps1    (renders theme cards)
    #
    #  Canonical token schema (REQUIRED -- every theme must provide):
    #    Name        Display name
    #    Vibe        Short description
    #    Bg          Window background
    #    Surface     Panel background
    #    Overlay     Input / inner-element background
    #    Border      Panel + input borders
    #    Hover       Hover state background
    #    Text        Body text (primary readable color)
    #    HeaderText  Column headers, reserved readability tone
    #    SubText     Subtitles, hint text (muted)
    #    Blue        Cool accent / "info" color (title, toggle-off label, etc.)
    #    Green       Primary action accent (Run button, toggle-on, checkmark)
    #    Red         Destructive / alert (Cancel button, status.Offline)
    #
    #  Optional richness fields (per-theme, opt-in):
    #    PipeStops         @('#top','#mid','#bottom')    Vertical gradient pipe
    #    TitleStops        @('#top','#bottom')           Vertical gradient title text
    #    RunStops          @('#left','#mid','#right')    Horizontal gradient Run button
    #    ProgressStops     @('#left','#mid','#right')    Horizontal gradient progress fill
    #    ToggleOnStops     @('#left','#mid','#right')    Horizontal gradient toggle-on track
    #    WindowBgGradient  @('#top','#bottom')           Subtle vertical window-bg gradient
    #    TitleStyle        'solid' | 'gradient' | 'split'   Default 'solid'
    #    TitleDimColor     '#hex'                        For split titles: dim "Invoke" color
    #    Info              '#hex'                        Secondary informational text
    #    CancelBg          '#hex'                        Pale tint for Cancel ghost bg
    #    CancelFg          '#hex'                        Cancel ghost foreground (default SubText)
    #    CancelBorder      '#hex'                        Cancel ghost border (default Border)
    #    PipeColor         '#hex'                        Pipe solid color (default Blue). Use when
    #                                                    the pipe's identity is a different accent
    #                                                    from Blue (e.g. Tanium red pipe).
    #    AccentText        '#hex'                        Text color on Green/Accent bg (default Bg).
    #                                                    Override for light themes where Bg is too
    #                                                    pale to read on a saturated accent.
    #    RunStyle          'solid' | 'ghost'             Default 'solid'
    #    RunGhostBg        '#hex'                        Deep-tinted bg for Run ghost style
    #    ToggleOnBg        '#hex'                        Toggle track on-state bg (default Green)
    #    ToggleOnBorder    '#hex'                        Toggle track on-state border (default
    #                                                    Green). Most themes use a lighter-shade
    #                                                    rim; set explicitly to preserve it.
    #    SubtitlePatch     'string'                      Subtitle text in Patch mode (default
    #                                                    'Patch Remediation'). For themes with
    #                                                    a distinct identity voice.
    #    SubtitleVersion   'string'                      Subtitle text in Version/Audit mode
    #                                                    (default 'Version Audit').
    #
    #  Defaults when an optional field is absent:
    #    No PipeStops     -> pipe is solid PipeColor (or Blue if PipeColor absent)
    #    No TitleStops and TitleStyle='solid' -> title is solid Blue
    #    No RunStops      -> Run button is solid Green (fg=AccentText)
    #    RunStyle='ghost' -> Run button bg=RunGhostBg, fg=Green, border=Green
    #    No ProgressStops -> progress fill is solid Green
    #    No ToggleOnStops -> toggle-on is solid Green
    #    No WindowBgGradient -> window is solid Bg
    #    No Info          -> secondary info text uses SubText
    #    No CancelBg      -> Cancel ghost bg uses Overlay
    #    No CancelFg      -> Cancel fg uses SubText (neutral-ghost default)
    #    No CancelBorder  -> Cancel border uses Border (neutral-ghost default)
    #    No AccentText    -> text on Run button uses Bg (dark on bright accent)
    # ========================================================================

    CobaltSlate = @{
        Name           = 'Cobalt Slate'
        Vibe           = 'Clean admin dashboard'
        Bg             = '#171C28'
        Surface        = '#1E2434'
        Overlay        = '#262E40'
        Border         = '#3E4860'
        Hover          = '#2A3244'
        Text           = '#E8ECF2'
        HeaderText     = '#C0CBE0'
        SubText        = '#8A98B8'
        Blue           = '#E8ECF2'
        Green          = '#E8ECF2'
        Red            = '#6878A0'
        Info           = '#6878A0'
        CancelFg       = '#6878A0'   # muted slate accent, border stays neutral
        ToggleOnBg     = '#5A8FCC'   # distinct medium-blue toggle (not white like Run)
        ToggleOnBorder = '#7AA5D8'   # lighter blue rim
    }

    CobaltSlateDay = @{
        Name         = 'Cobalt Slate Day'
        Vibe         = 'Light companion, cobalt on near-white'
        Bg           = '#EEF2F7'
        Surface      = '#FFFFFF'
        Overlay      = '#E6EBF2'
        Border       = '#C4CDD9'
        Hover        = '#DCE3ED'
        Text         = '#1C2333'
        HeaderText   = '#2E3A50'
        SubText      = '#6A7590'
        Blue           = '#2A4D80'
        Green          = '#2A4D80'
        Red            = '#8B6878'
        Info           = '#7A7F8A'
        AccentText     = '#FFFFFF'   # crisp white on cobalt (Bg too pale otherwise)
        CancelBg       = '#F2E8EB'   # pale rose tint
        CancelFg       = '#8B6878'   # rose accent ghost
        CancelBorder   = '#8B6878'
        ToggleOnBorder = '#4A6DA5'
    }

    TokyoNight = @{
        Name             = 'Tokyo Night'
        Vibe             = 'Pink + cyan on deep navy'
        Bg               = '#1A1B26'
        Surface          = '#24283B'
        Overlay          = '#2E3347'
        Border           = '#3B4261'
        Hover            = '#414868'
        Text             = '#C0CAF5'
        HeaderText       = '#C0CAF5'
        SubText          = '#9AA5CE'
        Blue             = '#7DCFFF'
        Green            = '#7DCFFF'
        Red              = '#F7768E'
        CancelFg         = '#F7768E'   # pink accent ghost
        CancelBorder     = '#F7768E'
        CancelBg         = '#4A2538'
        PipeStops        = @('#7DCFFF', '#BB9AF7', '#F7768E')
        TitleStops       = @('#7DCFFF', '#BB9AF7')
        RunStops         = @('#7DCFFF', '#BB9AF7', '#F7768E')
        ProgressStops    = @('#F7768E', '#BB9AF7', '#7DCFFF')
        ToggleOnStops    = @('#7DCFFF', '#BB9AF7', '#F7768E')
        WindowBgGradient = @('#1A1B26', '#12131C')
        TitleStyle       = 'gradient'
    }

    Meridian = @{
        Name       = 'Meridian'
        Vibe       = 'Naval charts, deep teal'
        Bg         = '#070A0B'
        Surface    = '#0C1315'
        Overlay    = '#121C1E'
        Border     = '#2C3A3D'
        Hover      = '#203032'
        Text       = '#DAE4E6'
        HeaderText = '#C4D4D8'
        SubText    = '#9AB0B4'
        Blue           = '#36B0A2'
        Green          = '#36B0A2'
        Red            = '#6A8888'
        RunStyle       = 'ghost'
        RunGhostBg     = '#0E2220'
        ToggleOnBorder = '#5FC8BC'
    }

    UltraDarkViolet = @{
        Name       = 'Ultra-Dark Violet'
        Vibe       = 'Near-black editorial'
        Bg         = '#0E0E0E'
        Surface    = '#161616'
        Overlay    = '#1E1E1E'
        Border     = '#404040'
        Hover      = '#302A32'
        Text       = '#E0E0E0'
        HeaderText = '#D0D0D0'
        SubText    = '#9C9C9C'
        Blue           = '#A06CD5'
        Green          = '#A06CD5'
        Red            = '#7A5A90'
        RunStyle       = 'ghost'
        RunGhostBg     = '#1E1628'
        ToggleOnBorder = '#B88CE0'
    }

    Quartz = @{
        Name       = 'Quartz'
        Vibe       = 'Pale mineral, ember accent'
        Bg         = '#C8C8D0'
        Surface    = '#D8D8E0'
        Overlay    = '#E2E2EA'
        Border     = '#B8B8C4'
        Hover      = '#D0D0DA'
        Text       = '#2A2A34'
        HeaderText = '#2A2A34'
        SubText    = '#4A4A58'
        Blue           = '#C85030'
        Green          = '#C85030'
        Red            = '#A84028'
        AccentText     = '#FFFFFF'   # Bg is light gray which fails contrast on coral
        ToggleOnBorder = '#A84028'   # darker coral shadow rim (not lighter)
    }

    DarkForest = @{
        Name             = 'Dark Forest'
        Vibe             = 'Canopy at twilight, copper accents'
        Bg               = '#0F171B'
        Surface          = '#1B2428'
        Overlay          = '#232D30'
        Border           = '#364248'
        Hover            = '#2E3A3C'
        Text             = '#D8DCD0'
        HeaderText       = '#C4CCB8'
        SubText          = '#9AA090'
        Blue             = '#A8C495'
        Green            = '#7FAC60'
        Red              = '#C08968'
        PipeColor        = '#C08968'   # copper pipe, not sage (Blue)
        CancelBg         = '#3A2F20'
        CancelFg         = '#C08968'   # copper accent ghost
        CancelBorder     = '#C08968'
        WindowBgGradient = @('#0F171B', '#0A1216')
    }

    NavyAnalytics = @{
        Name       = 'Navy Analytics'
        Vibe       = 'SCADA / instrument panel'
        Bg         = '#141820'
        Surface    = '#1A2028'
        Overlay    = '#222830'
        Border     = '#3C4458'
        Hover      = '#2A3038'
        Text       = '#C8CCD4'
        HeaderText = '#B8C4D4'
        SubText    = '#8898A8'
        Blue           = '#E0A030'
        Green          = '#E0A030'
        Red            = '#E06040'
        RunStyle       = 'ghost'
        RunGhostBg     = '#282018'
        ToggleOnBorder = '#F0C060'
    }

    Monochrome = @{
        Name          = 'Monochrome'
        Vibe          = 'Ink on paper'
        Bg            = '#121212'
        Surface       = '#1C1C1C'
        Overlay       = '#2A2A2A'
        Border        = '#4A4A4A'
        Hover         = '#353535'
        Text          = '#EAEAEA'
        HeaderText    = '#D4D4D4'
        SubText       = '#9C9C9C'
        Blue           = '#EAEAEA'
        Green          = '#EAEAEA'
        Red            = '#9C9C9C'
        TitleStyle     = 'split'
        TitleDimColor  = '#BEBEBE'
        ToggleOnBg     = '#9C9C9C'   # mid-gray track so the light knob stays visible
        ToggleOnBorder = '#FFFFFF'
    }

    TaniumInspired = @{
        Name       = 'Tanium-Inspired'
        Vibe       = 'Red branding, blue data'
        Bg         = '#0E1118'
        Surface    = '#161A24'
        Overlay    = '#1E2430'
        Border     = '#3A4258'
        Hover      = '#222838'
        Text       = '#E8EAF0'
        HeaderText = '#C0CBE0'
        SubText    = '#9098B4'
        Blue           = '#5098E0'
        Green          = '#5098E0'
        Red            = '#D03030'
        PipeColor      = '#D03030'   # Tanium brand-red pipe; Blue/Green is the data accent
        RunStyle       = 'ghost'
        RunGhostBg     = '#182848'
        ToggleOnBorder = '#78B4E8'
    }

    CarbonTeal = @{
        Name          = 'Carbon Teal'
        Vibe          = 'Carbon fiber, teal accents'
        Bg            = '#181818'
        Surface       = '#222222'
        Overlay       = '#2C2C2C'
        Border        = '#3C3C3C'
        Hover         = '#1C2C2C'
        Text          = '#E8E8E8'
        HeaderText    = '#D0D0D0'
        SubText       = '#9C9C9C'
        Blue           = '#00C8B4'
        Green          = '#00C8B4'
        Red            = '#009888'
        TitleStyle     = 'split'
        TitleDimColor  = '#B8B8B8'
        ToggleOnBorder = '#3FDCD0'
    }

    CyberCommand = @{
        Name       = 'CyberPunk Console'
        Vibe       = 'Neon cyan on pure black'
        Bg         = '#000000'
        Surface    = '#080810'
        Overlay    = '#101018'
        Border     = '#00E0FF'
        Hover      = '#001830'
        Text       = '#00E0FF'
        HeaderText = '#00E0FF'
        SubText    = '#3FB0D5'
        Blue       = '#00E0FF'
        Green      = '#00E0FF'
        Red        = '#00A0CC'
    }
}
