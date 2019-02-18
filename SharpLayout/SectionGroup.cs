using System.Collections.Generic;

namespace SharpLayout
{
    public class SectionGroup
    {
        public List<Section> Sections { get; } = new List<Section>();

        public Section Add(Section section)
        {
            Sections.Add(section);
            return section;
        }
    }
}