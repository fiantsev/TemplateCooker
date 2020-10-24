using FluentAssertions;
using TemplateCooker.Domain.Layout;
using Xunit;

namespace TemplateCookerUnitTests.Domain.Markers
{
    public class MarkerPositionTests
    {
        [Fact]
        public void MarkerPosition_должены_быть_равны_если_все_внутрение_свойства_одинаковые()
        {
            //assign
            var markerPosition1 = new SrcPosition(0, 0, 0);
            var markerPosition2 = new SrcPosition(0, 0, 0);

            //act
            //assert
            (markerPosition1 == markerPosition2).Should().BeTrue();
            object.Equals(markerPosition1, markerPosition2).Should().BeTrue();
            markerPosition1.Equals(markerPosition2).Should().BeTrue();
            markerPosition2.Equals(markerPosition1).Should().BeTrue();
        }

        [Fact]
        public void MarkerPosition_должены_быть_разные_если_внутрение_свойства_различаются()
        {
            //assign
            var markerPosition1 = new SrcPosition(1, 0, 0);
            var markerPosition2 = new SrcPosition(0, 0, 0);

            //act
            //assert
            (markerPosition1 == markerPosition2).Should().BeFalse();
            object.Equals(markerPosition1, markerPosition2).Should().BeFalse();
            markerPosition1.Equals(markerPosition2).Should().BeFalse();
            markerPosition2.Equals(markerPosition1).Should().BeFalse();
        }
    }
}
