using Xamarin.Forms;
using System.Collections.Generic;

namespace ProdLogProg
{
    public partial class ViewDataPage : ContentPage
    {
        public ViewDataPage(List<MainPage.ProductionData> dataList)
        {
            if (dataList == null || dataList.Count == 0)
            {
                DisplayAlert("Error", "No data to display.", "OK");
                return;
            }

            Title = "Excel Data";

            var listView = new ListView
            {
                ItemsSource = dataList,
                ItemTemplate = new DataTemplate(() =>
                {
                    var productLabel = new Label { FontAttributes = FontAttributes.Bold, TextColor = Color.Black };
                    productLabel.SetBinding(Label.TextProperty, "Product"); // Bind "Product" property

                    var timeLabel = new Label { TextColor = Color.Gray };
                    timeLabel.SetBinding(Label.TextProperty, "TimePerUnit"); // Bind "TimePerUnit" property

                    return new ViewCell
                    {
                        View = new StackLayout
                        {
                            Padding = new Thickness(10),
                            Children = { productLabel, timeLabel }
                        }
                    };
                })
            };

            var backButton = new Button
            {
                Text = "Back to Main Page",
                HorizontalOptions = LayoutOptions.Center
            };

            backButton.Clicked += async (sender, e) =>
            {
                await Navigation.PopAsync();
            };

            Content = new StackLayout
            {
                Children = { listView, backButton },
                Padding = new Thickness(10)
            };
        }
    }
}
