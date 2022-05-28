function Set-DocStyle {
    # Configure document options
    DocumentOption -EnableSectionNumbering -PageSize A4 -DefaultFont 'Arial' -MarginLeftAndRight 71 -MarginTopAndBottom 71

    # Configure Heading and Font Styles
    #Style -Name 'Title' -Size 24 -Color '000000' -Align Center
    #Style -Name 'Title 2' -Size 18 -Color '000000' -Align Center
    #Style -Name 'Title 3' -Size 12 -Color '000000' -Align Left
    Style -Name 'Heading 1' -Size 16 -Color 'B40000'
    Style -Name 'Heading 2' -Size 14 -Color '00852B'
    Style -Name 'Heading 3' -Size 12 -Color 'FAC80A'
    Style -Name 'Heading 4' -Size 11 -Color '1E5AA8'
    #Style -Name 'Heading 5' -Size 10 -Color '000000'
    Style -Name 'Normal' -Size 10 -Color '565656' -Default
    Style -Name 'Caption' -Size 10 -Color '565656' -Italic -Align Center
    Style -Name 'Header' -Size 10 -Color '565656' -Align Center
    Style -Name 'Footer' -Size 10 -Color '565656' -Align Center
    Style -Name 'TOC' -Size 16 -Color 'B40000'
    Style -Name 'TableDefaultHeading' -Size 10 -Color '000000' -BackgroundColor 'FFCF00'
    Style -Name 'TableDefaultRow' -Size 10 -Color '565656'

    # Configure Table Styles
    $TableDefaultProperties = @{
        Id = 'TableDefault'
        HeaderStyle = 'TableDefaultHeading'
        RowStyle = 'TableDefaultRow'
        BorderColor = '565656'
        Align = 'Left'
        CaptionStyle = 'Caption'
        CaptionLocation = 'Below'
        BorderWidth = 0.25
        PaddingTop = 1
        PaddingBottom = 1.5
        PaddingLeft = 2
        PaddingRight = 2
    }

    TableStyle @TableDefaultProperties -Default
}