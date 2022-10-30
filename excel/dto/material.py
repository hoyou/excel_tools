import attr
import attrs


@attr.s
class Material(object):
    material_name = attr.ib(type=str)
    factory_name = attr.ib(type=str, default="")
    material_brand_id = attr.ib(type=str, default="")
    time = attr.ib(type=str, default="")
    consume_num = attr.ib(type=str, default="")
    standard = attr.ib(type=str, default="")
    task_id = attr.ib(type=str, default="")
    brand_id = attr.ib(type=str, default="")
    flow_id = attr.ib(type=str, default="")
    material = attr.ib(type=str, default="")
    zhuding_flow_id = attr.ib(type=str, default="")


    @classmethod
    def map_2_convert(self, value_map: dict):
        if not dict:
            return None
        return Material(
            material_name=value_map.get("名称"),
            factory_name=value_map.get("厂家"),
            material_brand_id=value_map.get("牌号"),
            flow_id=value_map.get("批号"),
            consume_num=value_map.get("总用量/kg")
        )


if __name__ == '__main__':
    a1 = Material(material_name="test")
    print(a1)
